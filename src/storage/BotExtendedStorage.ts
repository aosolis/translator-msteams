// Copyright (c) Microsoft. All rights reserved.

import * as builder from "botbuilder";

/** Replacable storage system. */
export interface IBotExtendedStorage extends builder.IBotStorage {

    /** Reads in user data from storage based on AAD object id. */
    getUserDataByAadObjectIdAsync(aadObjectId: string): Promise<any>;

    /** Gets the AAD object id associated with the user data bag. */
    getAAdObjectId(userData: any): string;

    /** Sets the AAD object id associated with the user data bag. */
    setAAdObjectId(userData: any, aadObjectId: string): void;

}
