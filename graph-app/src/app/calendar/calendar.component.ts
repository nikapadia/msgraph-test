// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Component, OnInit } from '@angular/core';
import { parseISO } from 'date-fns';
import { endOfWeek, startOfWeek } from 'date-fns/esm';
import { zonedTimeToUtc } from 'date-fns-tz';
import { findIana } from 'windows-iana';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { AuthService } from '../auth.service';
import { GraphService } from '../graph.service';
import { AlertsService } from '../alerts.service';
import endOfMonth from 'date-fns/endOfMonth';

@Component({
  selector: 'app-calendar',
  templateUrl: './calendar.component.html',
  styleUrls: ['./calendar.component.css'],
})
export class CalendarComponent implements OnInit {
  public events?: MicrosoftGraph.Event[];

  constructor(
    private authService: AuthService,
    private graphService: GraphService,
    private alertsService: AlertsService
  ) {}

  async ngOnInit() {
    // Convert the user's timezone to IANA format
    const ianaName = findIana(this.authService.user?.timeZone ?? 'UTC');
    const timeZone =
      ianaName![0].valueOf() || this.authService.user?.timeZone || 'UTC';

    // Get midnight on the start of the current week in the user's timezone,
    // but in UTC. For example, for Pacific Standard Time, the time value would be
    // 07:00:00Z
    const now = new Date();
    const weekStart = zonedTimeToUtc(startOfWeek(now), timeZone);
    const weekEnd = zonedTimeToUtc(endOfMonth(now), timeZone);

    // weekEnd.setDate(weekEnd.getDate() + 7);

    this.events = await this.graphService.getCalendarView(
      weekStart.toISOString(),
      weekEnd.toISOString(),
      this.authService.user?.timeZone ?? 'UTC'
    );
  }

  async removeEvent(event: MicrosoftGraph.Event) {
      if (!this.authService.graphClient) {
          this.alertsService.addError('Graph client is not initialized.');
          return undefined;
      }
      try {
          await this.authService.graphClient
              .api(`/me/events/${event.id}`)
              .delete();
            this.events = this.events?.filter((e) => e.id !== event.id);
            return undefined;
      } catch (error) {
          this.alertsService.addError(
              'Could not delete event',
              JSON.stringify(error, null, 2)
          );
          return undefined;
      }
  }

  formatDateTimeTimeZone(
    dateTime: MicrosoftGraph.DateTimeTimeZone | undefined | null
  ): Date | undefined {
    if (dateTime == undefined || dateTime == null) {
      return undefined;
    }

    try {
      return parseISO(dateTime.dateTime!);
    } catch (error) {
      this.alertsService.addError(
        'DateTimeTimeZone conversion error',
        JSON.stringify(error)
      );
      return undefined;
    }
  }
}
