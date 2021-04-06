/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  Providers,
  ProviderState,
  MgtFastBaseComponent,
  prepScopes,
  MgtFastTemplatedComponent
} from '@microsoft/mgt-element';
import { customElement, attr, observable, html } from '@microsoft/fast-element';

import { getDayOfWeekString, getMonthString } from '../../utils/Utils';
import { getEventsPageIterator } from '../mgt-agenda/mgt-agenda.graph';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';

import { styles } from './mgt-agenda-fast-css';
import { template } from './mgt-agenda.template';

import '../mgt-person/mgt-person';
import '../../styles/style-helper';

/**
 * Web Component which represents events in a user or group calendar.
 *
 * @export
 * @class MgtAgenda
 * @extends {MgtTemplatedComponent}
 *
 * @fires eventClick - Fired when user click an event
 *
 * @cssprop --event-box-shadow - {String} Event box shadow color and size
 * @cssprop --event-margin - {String} Event margin
 * @cssprop --event-padding - {String} Event padding
 * @cssprop --event-background-color - {Color} Event background color
 * @cssprop --event-border - {String} Event border style
 * @cssprop --agenda-header-margin - {String} Agenda header margin size
 * @cssprop --agenda-header-font-size - {Length} Agenda header font size
 * @cssprop --agenda-header-color - {Color} Agenda header color
 * @cssprop --event-time-font-size - {Length} Event time font size
 * @cssprop --event-time-color - {Color} Event time color
 * @cssprop --event-subject-font-size - {Length} Event subject font size
 * @cssprop --event-subject-color - {Color} Event subject color
 * @cssprop --event-location-font-size - {Length} Event location font size
 * @cssprop --event-location-color - {Color} Event location color
 */
@customElement({
  name: 'mgt-fast-agenda',
  template,
  styles
})
export class MgtFastAgenda extends MgtFastTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * stores current date for initial calender selection in events.
   * @type {string}
   */
  @attr({ attribute: 'date' })
  public date: string;
  private dateChanged() {
    this.reload();
  }

  /**
   * determines if agenda events come from specific group
   * @type {string}
   */
  @attr({ attribute: 'group-id' })
  public groupId: string;
  private groupIdChanged() {
    this.reload();
  }

  /**
   * sets number of days until end date, 3 is the default
   * @type {number}
   */
  @attr({ attribute: 'days' })
  public days: number = 3;
  private daysChanged() {
    this.reload();
  }

  /**
   * allows developer to specify a different graph query that retrieves events
   * @type {string}
   */
  @attr({ attribute: 'event-query' })
  public eventQuery: string;
  private eventQueryChanged() {
    this.reload();
  }

  /**
   * array containing events from user agenda.
   * @type {MicrosoftGraph.Event[]}
   */
  @observable public events: MicrosoftGraph.Event[];

  /**
   * allows developer to define max number of events shown
   * @type {number}
   */
  @attr({ attribute: 'show-max' })
  public showMax: number;

  /**
   * allows developer to define agenda to group events by day.
   * @type {boolean}
   */
  @attr({ attribute: 'group-by-day', mode: 'boolean' })
  public groupByDay: boolean;

  /**
   * allows developer to specify preferred timezone that should be used for
   * retrieving events from Graph, eg. `Pacific Standard Time`. The preferred timezone for
   * the current user can be retrieved by calling `me/mailboxSettings` and
   * retrieving the value of the `timeZone` property.
   * @type {string}
   */
  @attr({ attribute: 'preferred-timezone' })
  public preferredTimezone: string;
  private preferredTimezoneChanged() {
    this.reload();
  }

  /**
   * determines width available for agenda component.
   * @type {boolean}
   */
  @observable private _isNarrow: boolean;

  constructor() {
    super();
    this.onResize = this.onResize.bind(this);
  }

  /**
   * Determines width available if resize is necessary, adds onResize event listener to window
   *
   * @memberof MgtAgenda
   */
  public connectedCallback() {
    this._isNarrow = this.offsetWidth < 600;
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
  }

  /**
   * Removes onResize event listener from window
   *
   * @memberof MgtAgenda
   */
  public disconnectedCallback() {
    window.removeEventListener('resize', this.onResize);
    super.disconnectedCallback();
  }

  // /**
  //  * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
  //  * Setting properties inside this method will not trigger the element to update
  //  *
  //  * @returns
  //  * @memberof MgtAgenda
  //  */
  // public render(): TemplateResult {
  //   // Loading
  //   if (!this.events && this.isLoadingState) {
  //     return this.renderLoading();
  //   }

  //   // No data
  //   if (!this.events || this.events.length === 0) {
  //     return this.renderNoData();
  //   }

  //   // Prep data
  //   const events = this.showMax && this.showMax > 0 ? this.events.slice(0, this.showMax) : this.events;

  //   // Default template
  //   const renderedTemplate = this.renderTemplate('default', { events });
  //   if (renderedTemplate) {
  //     return renderedTemplate;
  //   }

  //   // Update narrow state
  //   this._isNarrow = this.offsetWidth < 600;

  //   // Render list
  //   return html`
  //     <div dir=${this.direction} class="agenda${this._isNarrow ? ' narrow' : ''}${this.groupByDay ? ' grouped' : ''}">
  //       ${this.groupByDay ? this.renderGroups(events) : this.renderEvents(events)}
  //       ${this.isLoadingState ? this.renderLoading() : html``}
  //     </div>
  //   `;
  // }

  // /**
  //  * Render the loading state
  //  *
  //  * @protected
  //  * @returns
  //  * @memberof MgtAgenda
  //  */
  // protected renderLoading(): TemplateResult {
  //   return (
  //     this.renderTemplate('loading', null) ||
  //     html`
  //       <div class="event">
  //         <div class="event-time-container">
  //           <div class="event-time-loading loading-element"></div>
  //         </div>
  //         <div class="event-details-container">
  //           <div class="event-subject-loading loading-element"></div>
  //           <div class="event-location-container">
  //             <div class="event-location-icon-loading loading-element"></div>
  //             <div class="event-location-loading loading-element"></div>
  //           </div>
  //           <div class="event-location-container">
  //             <div class="event-attendee-loading loading-element"></div>
  //             <div class="event-attendee-loading loading-element"></div>
  //             <div class="event-attendee-loading loading-element"></div>
  //           </div>
  //         </div>
  //       </div>
  //     `
  //   );
  // }

  // /**
  //  * Render the no-data state.
  //  *
  //  * @protected
  //  * @returns {TemplateResult}
  //  * @memberof MgtAgenda
  //  */
  // protected renderNoData(): TemplateResult {
  //   return this.renderTemplate('no-data', null) || html``;
  // }

  // /**
  //  * Render the header for a group.
  //  * Only relevant for grouped Events.
  //  *
  //  * @protected
  //  * @param {Date} date
  //  * @returns
  //  * @memberof MgtAgenda
  //  */
  // protected renderHeader(header: string): TemplateResult {
  //   return (
  //     this.renderTemplate('header', { header }, 'header-' + header) ||
  //     html`
  //       <div class="header" aria-label="${header}">${header}</div>
  //     `
  //   );
  // }

  // /**
  //  * Render the events in groups, each with a header.
  //  *
  //  * @protected
  //  * @param {MicrosoftGraph.Event[]} events
  //  * @returns {TemplateResult}
  //  * @memberof MgtAgenda
  //  */
  // protected renderGroups(events: MicrosoftGraph.Event[]): TemplateResult {
  //   // Render list, grouped by day
  //   const grouped = {};

  //   events.forEach(event => {
  //     const header = this.getDateHeaderFromDateTimeString(event.start.dateTime);
  //     grouped[header] = grouped[header] || [];
  //     grouped[header].push(event);
  //   });

  //   return html`
  //     ${Object.keys(grouped).map(
  //       header =>
  //         html`
  //           <div class="group">${this.renderHeader(header)} ${this.renderEvents(grouped[header])}</div>
  //         `
  //     )}
  //   `;
  // }

  // /**
  //  * Render a list of events.
  //  *
  //  * @protected
  //  * @param {MicrosoftGraph.Event[]} events
  //  * @returns {TemplateResult}
  //  * @memberof MgtAgenda
  //  */
  // protected renderEvents(events: MicrosoftGraph.Event[]): TemplateResult {
  //   return html`
  //     <ul class="agenda-list">
  //       ${events.map(
  //         event =>
  //           html`
  //             <li @click=${() => this.eventClicked(event)}>
  //               ${this.renderTemplate('event', { event }, event.id) || this.renderEvent(event)}
  //             </li>
  //           `
  //       )}
  //     </ul>
  //   `;
  // }

  /**
   * Reloads the component with its current settings and potential new data
   *
   * @memberof MgtAgenda
   */
  public async reload() {
    this.clearState();
    this.requestStateUpdate(true);
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtAgenda
   */
  protected clearState(): void {
    this.events = null;
  }

  /**
   * Load state into the component
   *
   * @protected
   * @returns
   * @memberof MgtAgenda
   */
  protected async loadState() {
    if (this.events) {
      return;
    }

    const events = await this.loadEvents();
    if (events && events.length > 0) {
      this.events = events;
    }
  }

  private onResize() {
    this._isNarrow = this.offsetWidth < 600;
  }

  private eventClicked(event: MicrosoftGraph.Event) {
    this.$emit('eventClick', { event });
  }

  private async loadEvents(): Promise<MicrosoftGraph.Event[]> {
    const p = Providers.globalProvider;
    let events: MicrosoftGraph.Event[] = [];

    if (p && p.state === ProviderState.SignedIn) {
      const graph = p.graph.forComponent(this);

      if (this.eventQuery) {
        try {
          const tokens = this.eventQuery.split('|');
          let scope: string;
          let query: string;
          if (tokens.length > 1) {
            query = tokens[0].trim();
            scope = tokens[1].trim();
          } else {
            query = this.eventQuery;
          }

          let request = await graph.api(query);

          if (scope) {
            request = request.middlewareOptions(prepScopes(scope));
          }

          if (this.preferredTimezone) {
            request = request.header('Prefer', `outlook.timezone="${this.preferredTimezone}"`);
          }

          const results = await request.get();

          if (results && results.value) {
            events = results.value;
          }
          // tslint:disable-next-line: no-empty
        } catch (e) {}
      } else {
        const start = this.date ? new Date(this.date) : new Date();
        start.setHours(0, 0, 0, 0);
        const end = new Date(start.getTime());
        end.setDate(start.getDate() + this.days);
        try {
          const iterator = await getEventsPageIterator(graph, start, end, this.groupId, this.preferredTimezone);

          if (iterator && iterator.value) {
            events = iterator.value;

            while (iterator.hasNext) {
              await iterator.next();
              events = iterator.value;
            }
          }
        } catch (error) {
          // noop - possible error with graph
        }
      }
    }

    return events;
  }
}
