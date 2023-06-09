// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Injectable } from '@angular/core';

import { AuthService } from './auth.service';

@Injectable({
    providedIn: 'root',
})
export class GraphService {
    constructor(private authService: AuthService) {}

    // This is a hardcoded event creation function that is used for testing purposes
    async createTestEvent(): Promise<void> {
        if (!this.authService.graphClient) {
            console.log("Can't create event, no graph client");
            return undefined;
        }

        const testEvent = {
            subject: "Let's go for lunch",
            start: {
                dateTime: '2023-06-14T12:00:00',
                timeZone: 'Eastern Standard Time',
            },
            end: {
                dateTime: '2023-06-14T14:00:00',
                timeZone: 'Eastern Standard Time',
            },
        };

        try {
            await this.authService.graphClient
                .api('/me/events')
                .post(testEvent);
        } catch (error) {
            throw Error(JSON.stringify(error, null, 2));
        }
    }
}
