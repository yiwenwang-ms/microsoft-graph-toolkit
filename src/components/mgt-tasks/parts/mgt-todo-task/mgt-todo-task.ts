import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { OutlookTask } from '@microsoft/microsoft-graph-types-beta';
import { css, customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { Providers } from '../../../../Providers';
import { ProviderState } from '../../../../providers/IProvider';
import { getShortDateString } from '../../../../utils/Utils';
import '../../../common/mgt-button/mgt-button';
import '../../../common/mgt-context-menu/mgt-context-menu';
import { ContextMenuOption, MgtContextMenu } from '../../../common/mgt-context-menu/mgt-context-menu';
import '../../../common/mgt-icon/mgt-icon';
import { MgtTemplatedComponent } from '../../../templatedComponent';
import { ToDoTaskSource } from '../../taskSource/TodoTaskSource';
import { styles } from './mgt-todo-task-css';
/**
 *
 *
 * @export
 * @class MgtToDoTask
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-todo-task')
export class MgtToDoTask extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   *
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({
    attribute: 'task',
    type: Object
  })
  public task: OutlookTask;

  @property({ attribute: 'folder-name', type: String }) public folderName: string;

  @property({ attribute: 'read-only', type: Boolean }) public readOnly: boolean;

  @property({ attribute: false }) public taskSource: ToDoTaskSource;

  // assignment to this property will re-render the component
  @property({ attribute: false }) private _me: MicrosoftGraph.User;

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return this.renderTask();
  }

  private renderTask() {
    if (!this.task) {
      return null;
    }

    const task = this.task;

    const completed = task.status === 'completed';
    const dueDate = task.dueDateTime && new Date(task.dueDateTime.dateTime + 'Z');

    const isLoading = false; // TODO

    const context = { task: { ...task, folderTitle: this.folderName } };
    const taskTemplate = this.renderTemplate('task', context, task.id);
    if (taskTemplate) {
      return taskTemplate;
    }

    let taskDetails = this.renderTemplate('task-details', context, `task-details-${task.id}`);

    if (!taskDetails && dueDate) {
      taskDetails = html`
        <div class="TaskDetails">
          <div class="TaskDue">
            <span>Due ${getShortDateString(dueDate)}</span>
          </div>
        </div>
      `;
    }

    const contextMenuOptions: ContextMenuOption[] = [
      {
        icon: 'Delete',
        key: 'delete',
        onClick: () => this.removeTask(),
        text: 'Delete'
      }
    ];

    const taskOptions = this.readOnly
      ? null
      : html`
          <div class="TaskOptions">
            <mgt-context-menu .options="${contextMenuOptions}" @click=${e => e.stopPropagation()}>
              <mgt-button icon="More" @click=${this.menuClicked}></mgt-button>
            </mgt-context-menu>
          </div>
        `;

    const classes = {
      Complete: completed,
      ReadOnly: this.readOnly,
      Task: true,
      WithDetails: !!taskDetails
    };

    return html`
      <div class=${classMap(classes)}>
        <div class="TaskContent" @click="${() => this.handleTaskClick(task)}">
          <mgt-button
            class="TaskCheckButton"
            @click="${e => {
              this.toggleTask();
              e.stopPropagation();
            }}"
          >
            <span class="TaskCheck">
              <mgt-icon name="CheckMark"></mgt-icon>
            </span>
          </mgt-button>
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${task.subject}
            </div>
            ${taskDetails}
          </div>
          ${taskOptions}
        </div>
      </div>
    `;
  }

  private menuClicked() {
    const contextMenu = this.renderRoot.querySelector('mgt-context-menu') as MgtContextMenu;
    if (contextMenu) {
      contextMenu.isOpen = !contextMenu.isOpen;
    }
  }

  private async removeTask() {
    // todo
    if (this.taskSource) {
      this.task = await this.taskSource.removeTodoTask(this.task);
    }
  }

  private handleTaskClick(task) {
    // todo
    console.log('clicked', task);
  }

  private async toggleTask() {
    if (!this.readOnly) {
      if (this.task.status !== 'completed') {
        this.completeTask();
      } else {
        this.uncompleteTask();
      }
    }
  }

  private async completeTask() {
    // todo
    if (this.taskSource) {
      this.task = await this.taskSource.completeTask(this.task);
    }
  }

  private async uncompleteTask() {
    // todo
    if (this.taskSource) {
      this.task = await this.taskSource.incompleteTask(this.task);
    }
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    // TODO: load data from the graph
  }
}
