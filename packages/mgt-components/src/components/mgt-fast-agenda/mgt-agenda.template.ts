import { html, when, repeat } from '@microsoft/fast-element';
import { MgtFastAgenda } from './mgt-fast-agenda';

import './mgt-agenda-event/mgt-agenda-event';

export const template = html<MgtFastAgenda>`
  ${when(
    x => !x.events && x.isLoadingState,
    html`
      <div class="event">
        <div class="event-time-container">
          <div class="event-time-loading loading-element"></div>
        </div>
        <div class="event-details-container">
          <div class="event-subject-loading loading-element"></div>
          <div class="event-location-container">
            <div class="event-location-icon-loading loading-element"></div>
            <div class="event-location-loading loading-element"></div>
          </div>
          <div class="event-location-container">
            <div class="event-attendee-loading loading-element"></div>
            <div class="event-attendee-loading loading-element"></div>
            <div class="event-attendee-loading loading-element"></div>
          </div>
        </div>
      </div>
    `
  )}
  ${when(
    x => (!x.events || x.events.length === 0) && !x.isLoadingState,
    html`
      <div>TODO: no data</div>
    `
  )}
  ${when(
    x => x.events,
    html<MgtFastAgenda>`
      <div dir=${x => x.direction} class="agenda ${x => (x.groupByDay ? ' grouped' : '')}">
        <ul class="agenda-list">
          ${repeat(
            x => x.events,
            html<microsoftgraph.Event>`
              <div @click=${(x, c) => c.parent.eventClicked(x)}>
                <mgt-fast-agenda-event :event=${x => x}></mgt-fast-agenda-event>
              </div>
            `
          )}
        </ul>
      </div>
    `
  )}
`;
