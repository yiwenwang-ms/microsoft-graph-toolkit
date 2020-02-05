import { OutlookTask } from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, property } from 'lit-element';
import '../../../common/mgt-button/mgt-button';
import { MgtTemplatedComponent } from '../../../templatedComponent';
import { styles } from './mgt-todo-new-task-css';

/**
 *
 *
 * @export
 * @class mgt-todo-new-task
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-todo-new-task')
export class MgtNewToDoTask extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  @property({ attribute: false }) private taskName: string;
  @property({ attribute: false }) private taskDueDate: Date;
  @property({ attribute: false }) private loading: boolean;
  @property({ attribute: false }) private isDetailsVisible: boolean;

  constructor() {
    super();

    this.taskName = '';
  }

  protected toggleDetails() {
    this.isDetailsVisible = !this.isDetailsVisible;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const taskTitle = html`
      <input
        type="text"
        placeholder="Add a task"
        .value="${this.taskName}"
        label="new-taskName-input"
        aria-label="new-taskName-input"
        role="input"
        @input="${(e: Event) => {
          this.taskName = (e.target as HTMLInputElement).value;
        }}"
        @keyup=${e => {
          if (e.key === 'Enter') {
            this.addTask();
          }
        }}
      />
    `;

    return html`
      <div class="Task NewTask">
        <mgt-button type="primary" icon="Add" @click="${this.addTask}" .disabled=${this.taskName === ''}></mgt-button>
        <div class="TaskContent">
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${taskTitle}
            </div>
            ${this.renderDetails()}
          </div>
        </div>
        <mgt-button
          @click="${this.toggleDetails}"
          icon="${this.isDetailsVisible ? 'ChevronUp' : 'ChevronDown'}"
        ></mgt-button>
        <mgt-button icon="Cancel" @click="${this.cancel}"> </mgt-button>
      </div>
    `;
  }

  private renderDetails() {
    if (this.isDetailsVisible) {
      return html`
        <div class="TaskDetails">
          <div class="NewTaskDue">
            <input
              type="date"
              label="new-taskDate-input"
              aria-label="new-taskDate-input"
              role="input"
              .value="${this.dateToInputValue(this.taskDueDate)}"
              @change="${(e: Event) => {
                const value = (e.target as HTMLInputElement).value;
                if (value) {
                  this.taskDueDate = new Date(value + 'T17:00');
                } else {
                  this.taskDueDate = null;
                }
              }}"
            />
          </div>
        </div>
      `;
    }
  }

  private cancel() {
    this.dispatchEvent(new Event('cancel'));
  }

  private addTask() {
    if (this.taskName === '') {
      return;
    }

    const task: OutlookTask = {
      subject: this.taskName
    };

    if (this.taskDueDate) {
      task.dueDateTime = {
        dateTime: this.taskDueDate.toISOString(),
        timeZone: 'UTC'
      };
    }

    this.dispatchEvent(new CustomEvent('add', { detail: { task } }));

    console.log('add task', task);
  }

  private dateToInputValue(date: Date) {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }
}
