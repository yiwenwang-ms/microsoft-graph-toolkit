import { html, when, repeat } from '@microsoft/fast-element';
import { MgtFastAgendaEvent } from './mgt-agenda-event';

import '../../mgt-people/mgt-people';
import { getFastSvg, SvgIcon } from '../../../utils/SvgFastHelper';

export const template = html<MgtFastAgendaEvent>`
  <div class="event">
    <div class="event-time-container">
      <div class="event-time" aria-label="${x => x.eventTimeString}">${x => x.eventTimeString}</div>
    </div>
    <div class="event-details-container">
      <div class="event-subject">${x => x.event.subject}</div>
      ${when(
        x => x.event.location.displayName,
        html`
          <div class="event-location-container">
            <div class="event-location-icon">${getFastSvg(SvgIcon.OfficeLocation)}</div>
            <div class="event-location" aria-label="${x => x.event.location.displayName}">
              ${x => x.event.location.displayName}
            </div>
          </div>
        `
      )}
      ${when(
        x => x.event.attendees.length,
        html<MgtFastAgendaEvent>`
          <mgt-people
            class="event-attendees"
            :peopleQueries=${x => x.event.attendees.map(attendee => attendee.emailAddress.address)}
          ></mgt-people>
        `
      )}
    </div>
    <div class="event-other-container"></div>
  </div>
`;

// <div class="event-duration">${x => x.eventDuration}</div>
