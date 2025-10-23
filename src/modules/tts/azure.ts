import { getPref, setPref } from "../utils/prefs";
import { notifyGeneric } from "../utils/notify";
import { getString } from "../utils/locale";

// Azure Voice API response type
interface AzureVoice {
    Locale: string;
    ShortName: string;
    SecondaryLocaleList?: string[];
}

// Audio player class for handling Ogg/Opus playback
class AudioPlayer {
    private audioElement: HTMLAudioElement | null = null;
    private audioQueue: Uint8Array[] = [];
    private isPlaying: boolean = false;
    private isPaused: boolean = false;
    private isInitialized: boolean = false;

    public async initialize(): Promise<void> {
        if (this.isInitialized) {
            return;
        }

        this.audioElement = new window.Audio();
        this.audioElement.autoplay = false;

        // For Ogg/Opus, we can use direct blob URLs instead of MediaSource
        // since Firefox natively supports Ogg/Opus
        this.isInitialized = true;
    }

    public async queueAudioChunk(audioData: Uint8Array): Promise<void> {
        if (!this.isInitialized) {
            await this.initialize();
        }

        // Simply queue the chunk, will play all at turn.end
        this.audioQueue.push(audioData);
        ztoolkit.log(`Received audio chunk: ${audioData.length} bytes, queue size: ${this.audioQueue.reduce((sum, chunk) => sum + chunk.length, 0)} bytes`);
    }

    public async playRemaining(): Promise<void> {
        // Play any remaining audio in queue, regardless of size
        const queuedSize = this.audioQueue.reduce((sum, chunk) => sum + chunk.length, 0);
        ztoolkit.log(`playRemaining: queue size=${queuedSize} bytes, isPlaying=${this.isPlaying}`);
        if (!this.isPlaying && !this.isPaused && this.audioQueue.length > 0) {
            await this.playBatch();
        }
    }

    public pause(): void {
        if (this.audioElement && this.isPlaying) {
            this.audioElement.pause();
            this.isPaused = true;
            addon.data.tts.state = "paused";
        }
    }

    public resume(): void {
        if (this.audioElement && this.isPaused) {
            this.audioElement.play();
            this.isPaused = false;
            addon.data.tts.state = "playing";
        }
    }

    public stop(): void {
        this.audioQueue = [];
        this.isPlaying = false;
        this.isPaused = false;

        if (this.audioElement) {
            // Pause if playing
            if (!this.audioElement.paused) {
                this.audioElement.pause();
            }

            // Always clear src to prevent "Invalid URI" errors
            if (this.audioElement.src) {
                // Revoke blob URL if it's a blob
                if (this.audioElement.src.startsWith('blob:')) {
                    URL.revokeObjectURL(this.audioElement.src);
                }
                this.audioElement.removeAttribute('src');
                this.audioElement.load(); // Reset the element
            }
        }

        addon.data.tts.state = "idle";
    }

    public dispose(): void {
        this.stop();
        this.audioElement = null;
        this.isInitialized = false;
    }

    private async playBatch(): Promise<void> {
        if (this.audioQueue.length === 0 || this.isPaused || this.isPlaying) {
            return;
        }

        this.isPlaying = true;

        // Concatenate all queued audio chunks
        const totalLength = this.audioQueue.reduce((sum, chunk) => sum + chunk.length, 0);
        ztoolkit.log(`playBatch() starting: ${totalLength} bytes from ${this.audioQueue.length} chunks`);
        const concatenated = new Uint8Array(totalLength);
        let offset = 0;
        for (const chunk of this.audioQueue) {
            concatenated.set(chunk, offset);
            offset += chunk.length;
        }
        this.audioQueue = [];

        // Create blob and play
        const blob = new Blob([concatenated], { type: 'audio/ogg; codecs=opus' });
        const url = URL.createObjectURL(blob);

        if (this.audioElement) {
            this.audioElement.src = url;

            const playPromise = this.audioElement.play();

            if (playPromise !== undefined) {
                playPromise.catch((error) => {
                    ztoolkit.log(`Audio playback error: ${error}`);
                    this.isPlaying = false;
                    URL.revokeObjectURL(url);
                });
            }

            // When audio finishes, set state to idle
            this.audioElement.onended = () => {
                URL.revokeObjectURL(url);
                this.isPlaying = false;
                addon.data.tts.state = "idle";
            };
        }
    }
}

// WebSocket v2 connection manager for Azure Speech
class AzureStreamingSynthesizer {
    private ws: WebSocket | null = null;
    private requestId: string = '';
    private audioPlayer: AudioPlayer;
    private isConnected: boolean = false;
    private isStopped: boolean = false;
    private turnStartResolve: (() => void) | null = null;
    private turnStartReject: ((reason?: unknown) => void) | null = null;

    constructor() {
        this.audioPlayer = new AudioPlayer();
    }

    public async connect(): Promise<void> {
        if (this.isConnected && this.ws && this.ws.readyState === window.WebSocket.OPEN) {
            return;
        }

        const { key: subscriptionKey, region } = getAzureConfig();

        if (!subscriptionKey || !region) {
            throw new Error("auth-failed");
        }

        this.requestId = this.generateGuid();
        const connectionId = this.generateGuid();

        // Build WebSocket URL
        let wsUrl = `wss://${region}.tts.speech.microsoft.com/cognitiveservices/websocket/v2`;
        wsUrl += `?ConnectionId=${connectionId}`;
        wsUrl += `&X-ConnectionId=${connectionId}`;

        // Add subscription key to URL if provided
        if (subscriptionKey) {
            wsUrl += `&Ocp-Apim-Subscription-Key=${encodeURIComponent(subscriptionKey)}`;
        }

        return new Promise((resolve, reject) => {
            try {
                this.ws = new window.WebSocket(wsUrl);
                this.ws.binaryType = 'arraybuffer';

                this.ws.onopen = () => {
                    this.isConnected = true;
                    resolve();
                };

                this.ws.onmessage = async (event) => {
                    await this.handleMessage(event.data);
                };

                this.ws.onerror = (error) => {
                    ztoolkit.log(`WebSocket error: ${error}`);
                    if (!this.isConnected) {
                        reject(new Error("connection-failed"));
                    }
                };

                this.ws.onclose = (event) => {
                    const wasConnected = this.isConnected;
                    this.isConnected = false;
                    // Only notify if connection was previously established and then dropped
                    // Connection failures are handled by speak().catch()
                    if (event.code !== 1000 && !this.isStopped && wasConnected) {
                        ztoolkit.log(`WebSocket closed unexpectedly: ${event.code} ${event.reason}`);
                        notifyGeneric(
                            [getString("popup-engineErrorTitle", { args: { engine: "azure" } }),
                             getString("popup-engineErrorCause", { args: { engine: "azure", cause: "connection-closed" } })],
                            "error"
                        );
                    }
                };

            } catch (error) {
                reject(error);
            }
        });
    }

    public async speak(text: string): Promise<void> {
        this.isStopped = false;

        // Stop any previous playback
        if (this.audioPlayer) {
            this.audioPlayer.stop();
        }

        // Generate new requestId for each speak request
        this.requestId = this.generateGuid();

        await this.connect();
        await this.audioPlayer.initialize();

        // Send speech.config with empty object
        const configMessage = this.buildMessage('speech.config', {});

        this.ws?.send(configMessage);

        // Send synthesis.context
        const languageId = getPref("azure.language") as string || "en-US";
        const voiceName = getPref("azure.voice") as string || "en-US-AriaNeural";

        const contextMessage = this.buildMessage('synthesis.context', {
            synthesis: {
                audio: {
                    metadataOptions: {
                        sentenceBoundaryEnabled: false,
                        wordBoundaryEnabled: false,
                        visemeEnabled: false,
                        bookmarkEnabled: false,
                        punctuationBoundaryEnabled: false,
                        sessionEndEnabled: true
                    },
                    outputFormat: 'ogg-24khz-16bit-mono-opus'
                },
                language: {
                    autoDetection: false
                },
                input: {
                    bidirectionalStreamingMode: true,
                    voiceName: voiceName,
                    language: languageId
                }
            }
        });

        // Reset turn.start state and create Promise to wait for it
        const turnStartPromise = new Promise<void>((resolve, reject) => {
            this.turnStartResolve = resolve;
            this.turnStartReject = reject;
        });

        this.ws?.send(contextMessage);

        // Wait for turn.start response with timeout
        try {
            await Promise.race([
                turnStartPromise,
                new Promise<void>((_, reject) =>
                    setTimeout(() => reject(new Error("Timeout waiting for turn.start")), 5000)
                )
            ]);

            // Check if cancelled during wait
            if (this.isStopped) {
                ztoolkit.log("Synthesis cancelled while waiting for turn.start");
                return;
            }
        } catch (error) {
            ztoolkit.log(`Error waiting for turn.start: ${error}`);
            throw error;
        }

        // Send all text at once
        ztoolkit.log(`Sending full text: ${text.length} characters`);
        addon.data.tts.state = "playing";

        const textMessage = this.buildMessage('text.piece', text, 'text/plain');
        this.ws?.send(textMessage);

        // Immediately send text.end to signal completion
        const endMessage = this.buildMessage('text.end', '', 'text/plain');
        this.ws?.send(endMessage);

        ztoolkit.log(`Text sent, waiting for audio synthesis to complete`)
    }

    public stop(): void {
        this.isStopped = true;

        // Interrupt waiting for turn.start if in progress
        if (this.turnStartReject) {
            this.turnStartReject(new Error("Cancelled by user"));
            this.turnStartReject = null;
            this.turnStartResolve = null;
        }

        this.audioPlayer.stop();
        addon.data.tts.state = "idle";
    }

    public pause(): void {
        this.audioPlayer.pause();
    }

    public resume(): void {
        this.audioPlayer.resume();
    }

    public disconnect(): void {
        if (this.ws) {
            this.ws.close(1000, "Normal closure");
            this.ws = null;
        }
        this.isConnected = false;
    }

    public dispose(): void {
        this.stop();
        this.disconnect();
        this.audioPlayer.dispose();
    }

    private generateGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
            const r = Math.random() * 16 | 0;
            const v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    private generateTimestamp(): string {
        return new Date().toISOString();
    }

    private buildMessage(path: string, body: string | Record<string, unknown>, contentType: string = 'application/json'): string {
        const timestamp = this.generateTimestamp();
        const headers = [
            `X-Timestamp:${timestamp}`,
            `X-RequestId:${this.requestId}`,
            `Path:${path}`,
            `Content-Type:${contentType}`,
            '',
            ''
        ];

        // For text/plain messages, body is already a string; for JSON messages, stringify it
        const content = contentType === 'text/plain' ? body : JSON.stringify(body);
        return headers.join('\r\n') + content;
    }

    private async handleMessage(data: ArrayBuffer | string): Promise<void> {
        if (typeof data === 'string') {
            // Parse headers from text message
            const headerEnd = data.indexOf('\r\n\r\n');
            if (headerEnd !== -1) {
                const headerText = data.substring(0, headerEnd);
                const lines = headerText.split('\r\n');
                let path = '';

                for (const line of lines) {
                    if (line.startsWith('Path:')) {
                        path = line.substring(5).trim();
                        break;
                    }
                }

                // Only log important messages
                if (path === 'turn.start') {
                    ztoolkit.log('turn.start received');
                    if (this.turnStartResolve) {
                        this.turnStartResolve();
                        this.turnStartResolve = null;
                        this.turnStartReject = null;
                    }
                } else if (path === 'turn.end') {
                    ztoolkit.log(`turn.end received`);
                    // Play all accumulated audio
                    await this.audioPlayer.playRemaining();
                } else if (path === 'response') {
                    // Ignore response messages
                } else {
                    // Log unknown message types
                    ztoolkit.log(`Unknown text message: Path=${path}`);
                }
            }
            return;
        }

        // Binary message - Azure uses 2-byte length prefix for headers
        const view = new Uint8Array(data);

        if (view.length < 2) {
            ztoolkit.log(`Binary message too short: ${view.length} bytes`);
            return;
        }

        // Read header length from first 2 bytes (big-endian)
        const headerLength = (view[0] << 8) | view[1];

        if (view.length < 2 + headerLength) {
            ztoolkit.log(`Binary message incomplete: expected ${2 + headerLength} bytes, got ${view.length}`);
            return;
        }

        // Extract header text
        const headerBytes = view.slice(2, 2 + headerLength);
        const headerText = new TextDecoder('utf-8').decode(headerBytes);
        const headers = this.parseHeaders(headerText);

        if (headers['Path'] === 'audio') {
            // Audio chunk - extract audio data after header
            const audioData = view.slice(2 + headerLength);

            if (audioData.length > 0) {
                ztoolkit.log(`Received audio chunk: ${audioData.length} bytes`);
                try {
                    await this.audioPlayer.queueAudioChunk(audioData);
                } catch (error) {
                    ztoolkit.log(`Error queueing audio chunk: ${error}`);
                }
            }
        } else if (headers['Path'] === 'response') {
            // Response metadata
            const bodyStart = 2 + headerLength;
            if (bodyStart < view.length) {
                const bodyText = new TextDecoder('utf-8').decode(view.slice(bodyStart));
                ztoolkit.log(`Response: ${bodyText}`);
            }
        }
    }

    private parseHeaders(headerText: string): { [key: string]: string } {
        const headers: { [key: string]: string } = {};
        const lines = headerText.split('\r\n');

        for (const line of lines) {
            const colonIndex = line.indexOf(':');
            if (colonIndex > 0) {
                const key = line.substring(0, colonIndex).trim();
                const value = line.substring(colonIndex + 1).trim();
                headers[key] = value;
            }
        }

        return headers;
    }
}

// Singleton instance management
let synthesizer: AzureStreamingSynthesizer | null = null;

function getSynthesizer(): AzureStreamingSynthesizer {
    if (!synthesizer) {
        synthesizer = new AzureStreamingSynthesizer();
    }
    return synthesizer;
}

function setDefaultPrefs(): void {
    if (!getPref("azure.subscriptionKey")) {
        setPref("azure.subscriptionKey", "");
    }

    if (!getPref("azure.region")) {
        setPref("azure.region", "");
    }

    if (!getPref("azure.language")) {
        setPref("azure.language", "en-US");
    }

    // No default voice - user must select after fetching from API
}

async function initEngine(): Promise<void> {
    // Azure engine initialization always succeeds
    // Actual validation happens when user tries to speak
    // This allows users to configure the engine after installation
    return Promise.resolve();
}

// Get Azure configuration from environment variables and preferences
// Preferences override environment variables
function getAzureConfig(): { key: string; region: string } {
    let subscriptionKey = "";
    let region = "";

    // Try environment variables first
    try {
        // @ts-ignore - nsIEnvironment not in type definitions
        const env = Components.classes["@mozilla.org/process/environment;1"]
            .getService(Components.interfaces.nsIEnvironment);
        if (env.exists("AZURE_SPEECH_KEY")) {
            subscriptionKey = env.get("AZURE_SPEECH_KEY");
        }
        if (env.exists("AZURE_SPEECH_REGION")) {
            region = env.get("AZURE_SPEECH_REGION");
        }
    } catch (error) {
        ztoolkit.log(`Failed to read Azure environment variables: ${error}`);
    }

    // Preferences override environment variables
    const prefKey = (getPref("azure.subscriptionKey") as string || "").trim();
    if (prefKey) {
        subscriptionKey = prefKey;
    }

    const prefRegion = (getPref("azure.region") as string || "").trim();
    if (prefRegion) {
        region = prefRegion;
    }

    return {
        key: subscriptionKey.trim(),
        region: region.trim()
    };
}

// Exported functions matching webSpeech.ts pattern
function speak(text: string): void {
    const synth = getSynthesizer();

    synth.speak(text).catch((error) => {
        ztoolkit.log(`Azure TTS error: ${error}`);

        let errorKey = "other";
        if (error.message === "auth-failed") {
            errorKey = "auth-failed";
        } else if (error.message === "connection-failed") {
            errorKey = "connection-failed";
        }

        notifyGeneric(
            [getString("popup-engineErrorTitle", { args: { engine: "azure" } }),
             getString("popup-engineErrorCause", { args: { engine: "azure", cause: errorKey } })],
            "error"
        );

        addon.data.tts.state = "idle";
    });
}

function stop(): void {
    if (synthesizer) {
        synthesizer.stop();
    }
}

function pause(): void {
    if (synthesizer) {
        synthesizer.pause();
    }
}

function resume(): void {
    if (synthesizer) {
        synthesizer.resume();
    }
}

function resetConnection(): void {
    if (synthesizer) {
        synthesizer.disconnect();
    }
}

function dispose(): void {
    if (synthesizer) {
        synthesizer.dispose();
        synthesizer = null;
    }
}

// Extras for preferences UI
async function getAllVoices(): Promise<{ success: boolean; voices: AzureVoice[] }> {
    const { key: subscriptionKey, region } = getAzureConfig();

    if (!subscriptionKey || !region) {
        ztoolkit.log("No subscription key or region available");
        return { success: false, voices: [] };
    }

    try {
        const url = `https://${region}.tts.speech.microsoft.com/cognitiveservices/voices/list`;
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Ocp-Apim-Subscription-Key': subscriptionKey
            }
        });

        if (!response.ok) {
            ztoolkit.log(`Failed to fetch voices: ${response.status} ${response.statusText}`);
            return { success: false, voices: [] };
        }

        const voices = await response.json();

        if (!Array.isArray(voices)) {
            ztoolkit.log("Invalid response format from Azure voices API");
            return { success: false, voices: [] };
        }

        return { success: true, voices: voices };

    } catch (error) {
        ztoolkit.log(`Error fetching voices from Azure: ${error}\n  URL: https://${region}.tts.speech.microsoft.com/cognitiveservices/voices/list\n  Region: ${region}\n  Key length: ${subscriptionKey.length}`);
        return { success: false, voices: [] };
    }
}

function extractLanguages(voices: AzureVoice[]): string[] {
    const languageSet = new Set<string>();

    voices.forEach(voice => {
        if (voice.Locale) {
            languageSet.add(voice.Locale);
        }

        if (voice.SecondaryLocaleList && Array.isArray(voice.SecondaryLocaleList)) {
            voice.SecondaryLocaleList.forEach((lang: string) => {
                languageSet.add(lang);
            });
        }
    });

    return Array.from(languageSet).sort();
}

function filterVoicesByLanguage(voices: AzureVoice[], language: string): string[] {
    const filtered = voices.filter(voice => {
        if (voice.Locale === language) {
            return true;
        }

        if (voice.SecondaryLocaleList && Array.isArray(voice.SecondaryLocaleList)) {
            return voice.SecondaryLocaleList.includes(language);
        }

        return false;
    });

    return filtered
        .map(voice => voice.ShortName)
        .filter(name => name) // Filter out any undefined/null
        .sort();
}

export {
    speak,
    stop,
    pause,
    resume,
    resetConnection,
    dispose,
    setDefaultPrefs,
    initEngine,
    getAzureConfig,
    getAllVoices,
    extractLanguages,
    filterVoicesByLanguage
};
