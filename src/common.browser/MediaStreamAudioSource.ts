// Original work Copyright (c) Microsoft Corporation. All rights reserved.
// Modified work (c) Martin Stöber
// Licensed under the MIT license.

import {
    connectivity,
    ISpeechConfigAudioDevice,
    type
} from "../common.speech/Exports";
import {
    AudioSourceErrorEvent,
    AudioSourceEvent,
    AudioSourceInitializingEvent,
    AudioSourceOffEvent,
    AudioSourceReadyEvent,
    AudioStreamNodeAttachedEvent,
    AudioStreamNodeAttachingEvent,
    AudioStreamNodeDetachedEvent,
    AudioStreamNodeErrorEvent,
    ChunkedArrayBufferStream,
    createNoDashGuid,
    Events,
    EventSource,
    IAudioSource,
    IAudioStreamNode,
    IStringDictionary,
    Stream,
} from "../common/Exports";
import {
    AudioStreamFormat,
    AudioStreamFormatImpl,
} from "../sdk/Audio/AudioStreamFormat";
import { IRecorder } from "./IRecorder";

export const MediaStreamAudioWorkletSourceURLPropertyName = "MEDIASTREAM-WorkletSourceUrl";

export class MediaStreamAudioSource implements IAudioSource {

    private static readonly AUDIOFORMAT: AudioStreamFormatImpl = AudioStreamFormat.getDefaultInputFormat() as AudioStreamFormatImpl;

    private privStreams: IStringDictionary<Stream<ArrayBuffer>> = {};

    private privId: string;

    private privEvents: EventSource<AudioSourceEvent>;

    private privMediaStream: MediaStream;

    private privContext: AudioContext;

    private privOutputChunkSize: number;

    public constructor(
        private readonly privRecorder: IRecorder,
        private readonly mediaStream: MediaStream,
        audioSourceId?: string) {

        this.privOutputChunkSize = MediaStreamAudioSource.AUDIOFORMAT.avgBytesPerSec / 10;
        this.privId = audioSourceId ? audioSourceId : createNoDashGuid();
        this.privEvents = new EventSource<AudioSourceEvent>();
    }

    public get format(): Promise<AudioStreamFormatImpl> {
        return Promise.resolve(MediaStreamAudioSource.AUDIOFORMAT);
    }

    public get blob(): Promise<Blob> {
        return Promise.reject("Not implemented for MediaStream input");
    }

    public turnOn = (): Promise<void> => {
        this.onEvent(new AudioSourceInitializingEvent(this.privId)); // no stream id
        try {
            this.createAudioContext();
        } catch (error) {
            if (error instanceof Error) {
                const typedError: Error = error as Error;
                this.onEvent(new AudioSourceErrorEvent(this.privId, typedError.message));
                Promise.reject(typedError.name + ": " + typedError.message);
            } else {
                this.onEvent(new AudioSourceErrorEvent(this.privId, "Could not create AudioContext."));
                Promise.reject(error);
            }
        }

        this.privMediaStream = this.mediaStream;
        this.onEvent(new AudioSourceReadyEvent(this.privId));

        return;
    }

    public id = (): string => {
        return this.privId;
    }

    public attach = (audioNodeId: string): Promise<IAudioStreamNode> => {
        this.onEvent(new AudioStreamNodeAttachingEvent(this.privId, audioNodeId));

        return this.listen(audioNodeId).then<IAudioStreamNode>(
            (stream: Stream<ArrayBuffer>) => {
                this.onEvent(new AudioStreamNodeAttachedEvent(this.privId, audioNodeId));
                return {
                    detach: async () => {
                        stream.readEnded();
                        delete this.privStreams[audioNodeId];
                        this.onEvent(new AudioStreamNodeDetachedEvent(this.privId, audioNodeId));
                        return this.turnOff();
                    },
                    id: () => {
                        return audioNodeId;
                    },
                    read: () => {
                        return stream.read();
                    },
                };
            });
    }

    public detach = (audioNodeId: string): void => {
        if (audioNodeId && this.privStreams[audioNodeId]) {
            this.privStreams[audioNodeId].close();
            delete this.privStreams[audioNodeId];
            this.onEvent(new AudioStreamNodeDetachedEvent(this.privId, audioNodeId));
        }
    }

    public async turnOff(): Promise<void> {
        for (const streamId in this.privStreams) {
            if (streamId) {
                const stream = this.privStreams[streamId];
                if (stream) {
                    stream.close();
                }
            }
        }

        this.onEvent(new AudioSourceOffEvent(this.privId)); // no stream now

        await this.destroyAudioContext();

        return;
    }

    public get events(): EventSource<AudioSourceEvent> {
        return this.privEvents;
    }

    public get deviceInfo(): Promise<ISpeechConfigAudioDevice> {
        return Promise.resolve({
                bitspersample: MediaStreamAudioSource.AUDIOFORMAT.bitsPerSample,
                channelcount: MediaStreamAudioSource.AUDIOFORMAT.channels,
                connectivity: connectivity.Unknown,
                manufacturer: "Custom Speech SDK",
                model: "MediaStream",
                samplerate: MediaStreamAudioSource.AUDIOFORMAT.samplesPerSec,
                type: type.Stream,
        });
    }

    public setProperty(name: string, value: string): void {
        if (name === MediaStreamAudioWorkletSourceURLPropertyName) {
            this.privRecorder.setWorkletUrl(value);
        } else {
            throw new Error("Property '" + name + "' is not supported on MediaStream.");
        }
    }

    private listen = async (audioNodeId: string): Promise<Stream<ArrayBuffer>> => {
        await this.turnOn();
        const stream = new ChunkedArrayBufferStream(this.privOutputChunkSize, audioNodeId);
        this.privStreams[audioNodeId] = stream;
        try {
            this.privRecorder.record(this.privContext, this.privMediaStream, stream);
        } catch (error) {
            this.onEvent(new AudioStreamNodeErrorEvent(this.privId, audioNodeId, error));
            throw error;
        }
        const result: Stream<ArrayBuffer> = stream;
        return result;
    }

    private onEvent = (event: AudioSourceEvent): void => {
        this.privEvents.onEvent(event);
        Events.instance.onEvent(event);
    }

    private createAudioContext = (): void => {
        if (!!this.privContext) {
            return;
        }

        this.privContext = AudioStreamFormatImpl.getAudioContext(MediaStreamAudioSource.AUDIOFORMAT.samplesPerSec);
    }

    private async destroyAudioContext(): Promise<void> {
        if (!this.privContext) {
            return;
        }

        this.privRecorder.releaseMediaResources(this.privContext);

        // This pattern brought to you by a bug in the TypeScript compiler where it
        // confuses the ("close" in this.privContext) with this.privContext always being null as the alternate.
        // https://github.com/Microsoft/TypeScript/issues/11498
        let hasClose: boolean = false;
        if ("close" in this.privContext) {
            hasClose = true;
        }

        if (hasClose) {
            await this.privContext.close();
            this.privContext = null;
        } else if (null !== this.privContext && this.privContext.state === "running") {
            // Suspend actually takes a callback, but analogous to the
            // resume method, it'll be only fired if suspend is called
            // in a direct response to a user action. The later is not always
            // the case, as TurnOff is also called, when we receive an
            // end-of-speech message from the service. So, doing a best effort
            // fire-and-forget here.
            await this.privContext.suspend();
        }
    }
}
