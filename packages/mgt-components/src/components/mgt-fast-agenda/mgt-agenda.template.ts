import { html, when, repeat } from '@microsoft/fast-element';
import { MgtFastAgenda } from './mgt-fast-agenda';

import './mgt-agenda-event/mgt-agenda-event';

export const template = html<MgtFastAgenda>`
  ${when(x => !x.events && x.isLoadingState, x => x.getTemplate('loading', null, 'loading', null))}
  ${when(
    x => (!x.events || x.events.length === 0) && !x.isLoadingState,
    x =>
      x.getTemplate(
        'no-data',
        null,
        'no-data',
        html`
          <div>TODO: no data</div>
        `
      )
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
                ${(x, c) =>
                  c.parent.getTemplate(
                    'event',
                    { event: x },
                    x.id,
                    html`
                      <mgt-fast-agenda-event :event=${x => x}></mgt-fast-agenda-event>
                    `
                  )}
              </div>
            `
          )}
        </ul>
      </div>
    `
  )}
`;
