import * as assert from "assert";
import * as mongodb from "mongodb";
import * as winston from "winston";
import { IBotChannelStorageContext, IBotChannelStorageData, IBotChannelStorage } from "./BotChannelStorage";

/** Replacable storage system used by UniversalBot. */
export class MongoDbBotChannelStorage implements IBotChannelStorage {

    private initializePromise: Promise<void>;
    private mongoDb: mongodb.Db;
    private botStateCollection: mongodb.Collection;

    constructor(
        private collectionName: string,
        private connectionString: string) {
    }

    // Reads in data from storage
    public async getData(context: IBotChannelStorageContext): Promise<IBotChannelStorageData> {
        if (context.channelId) {
            await this.initialize();

            let filter = { "key": this.getChannelDataId(context) };
            let document = await this.botStateCollection.findOne(filter);
            if (document && document.data) {
                return { channelData: document.data };
            } else {
                return { };
            }
        } else {
            return { };
        }
    }

    // Writes out data from storage
    public async saveData(context: IBotChannelStorageContext, data: IBotChannelStorageData): Promise<void> {
        if (context.teamId && context.channelId && data.channelData) {
            await this.initialize();

            let id = this.getChannelDataId(context);
            let filter = { key: id };
            let document = {
                key: id,
                teamId: context.teamId,
                channelId: context.channelId,
                data: data.channelData,
                lastUpdate: new Date().valueOf(),
            };
            await this.botStateCollection.updateOne(filter, document, { upsert: true });
        }
    }

    // Returns a promise that is resolved when this instance is initialized
    private initialize(): Promise<void> {
        if (!this.initializePromise) {
            this.initializePromise = this.initializeWorker();
        }
        return this.initializePromise;
    }

    // Initialize this instance
    private async initializeWorker(): Promise<void> {
        if (!this.mongoDb) {
            try {
                this.mongoDb = await mongodb.MongoClient.connect(this.connectionString);
                this.botStateCollection = await this.mongoDb.collection(this.collectionName);

                // Set up indexes
                await this.botStateCollection.createIndex({ key: 1 });
                await this.botStateCollection.createIndex({ lastUpdate: 1 });
            } catch (e) {
                winston.error("Error initializing MongoDB channel storage", e);
                this.close();
                this.initializePromise = null;
            }
        }
    }

    // Close the connection to the database
    private close(): void {
        this.botStateCollection = null;
        if (this.mongoDb) {
            this.mongoDb.close();
            this.mongoDb = null;
        }
    }

    // Get id for channel data documents
    private getChannelDataId(context: IBotChannelStorageContext): string {
        assert(context.channelId);
        return `channel:${context.channelId}`;
    }
}
