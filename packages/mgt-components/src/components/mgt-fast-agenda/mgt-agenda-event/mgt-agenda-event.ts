import { MgtFastBaseComponent } from '@microsoft/mgt-element';
import { customElement, observable } from '@microsoft/fast-element';

import { template } from './mgt-agenda-event.template';
import { styles } from './mgt-agenda-event-fast-css';
import { getDayOfWeekString, getMonthString } from '../../../utils/Utils';

@customElement({
  name: 'mgt-fast-agenda-event',
  template,
  styles
})
export class MgtFastAgendaEvent extends MgtFastBaseComponent {
  @observable public event: microsoftgraph.Event;
  private eventChanged() {
    this.eventTimeString = this.getEventTimeString();
    this.eventDuration = this.getEventDuration();
  }

  @observable public eventTimeString: string;
  @observable public eventDuration: string;

  @observable public preferredTimezone: string;

  private getEventTimeString() {
    if (this.event.isAllDay) {
      return 'ALL DAY';
    }

    const start = this.prettyPrintTimeFromDateTime(new Date(this.event.start.dateTime));
    const end = this.prettyPrintTimeFromDateTime(new Date(this.event.end.dateTime));

    return `${start} - ${end}`;
  }

  private prettyPrintTimeFromDateTime(date: Date) {
    // If a preferred time zone was sent in the Graph request
    // times are already set correctly. Do not adjust
    if (!this.preferredTimezone) {
      // If no preferred time zone was specified, the times are in UTC
      // fall back to old behavior and adjust the times to the browser's
      // time zone
      date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    }

    let hours = date.getHours();
    const minutes = date.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const minutesStr = minutes < 10 ? '0' + minutes : minutes;
    return `${hours}:${minutesStr} ${ampm}`;
  }

  private getDateHeaderFromDateTimeString(dateTimeString: string) {
    const date = new Date(dateTimeString);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());

    const dayIndex = date.getDay();
    const monthIndex = date.getMonth();
    const day = date.getDate();
    const year = date.getFullYear();

    return `${getDayOfWeekString(dayIndex)}, ${getMonthString(monthIndex)} ${day}, ${year}`;
  }

  private getEventDuration() {
    let dtStart = new Date(this.event.start.dateTime);
    const dtEnd = new Date(this.event.end.dateTime);
    const dtNow = new Date();
    let result: string = '';

    if (dtNow > dtStart) {
      dtStart = dtNow;
    }

    const diff = dtEnd.getTime() - dtStart.getTime();
    const durationMinutes = Math.round(diff / 60000);

    if (durationMinutes > 1440 || this.event.isAllDay) {
      result = Math.ceil(durationMinutes / 1440) + 'd';
    } else if (durationMinutes > 60) {
      result = Math.round(durationMinutes / 60) + 'h';
      const leftoverMinutes = durationMinutes % 60;
      if (leftoverMinutes) {
        result += leftoverMinutes + 'm';
      }
    } else {
      result = durationMinutes + 'm';
    }

    return result;
  }
}
