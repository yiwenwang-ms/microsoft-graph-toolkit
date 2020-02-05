import { GraphRequest } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { Graph } from '../../../Graph';
import { EventDispatcher, EventHandler } from '../../../utils/EventDispatcher';
import { prepScopes } from '../../../utils/GraphHelpers';
import { BaseTaskSource } from './baseTaskSource';

/**
 * TODO
 *
 * @export
 * @class ToDoTaskSource
 */
export class ToDoTaskSource
  implements
    BaseTaskSource<
      MicrosoftGraphBeta.OutlookTask,
      MicrosoftGraphBeta.OutlookTaskFolder,
      MicrosoftGraphBeta.OutlookTaskGroup
    > {
  private readonly scopes = {
    read: 'Tasks.Read',
    write: 'Tasks.ReadWrite'
  };

  private readonly nextLinkString = '@odata.nextLink';

  private graph: Graph;
  private tasks: Record<string, { value: MicrosoftGraphBeta.OutlookTask[] }> = {};
  private folders: MicrosoftGraphBeta.OutlookTaskFolder[] = [];

  private taskUpdatedEventDispatcher = new EventDispatcher<Event>();

  constructor(graph: Graph) {
    this.graph = graph;
  }

  /**
   * Get all tasks
   *
   * @static
   * @memberof ToDoTaskSource
   */
  public async getTasks(folderId): Promise<MicrosoftGraphBeta.OutlookTask[]> {
    let tasks = this.tasks[folderId];
    if (!tasks) {
      tasks = await this.graph
        .api(`me/outlook/taskFolders/${folderId}/tasks`)
        .middlewareOptions(prepScopes(this.scopes.read))
        .version('beta')
        .get();

      this.tasks[folderId] = tasks;
    }

    return tasks && tasks.value;
  }

  public async getFolders(): Promise<MicrosoftGraphBeta.OutlookTaskFolder[]> {
    if (this.folders.length > 0) {
      return this.folders;
    }

    const request = await this.graph
      .api('me/outlook/taskFolders')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    this.folders = request && request.value;
    return this.folders;
  }

  public hasMoreForFolder(folderId: string) {
    const tasks = this.tasks[folderId];
    if (!tasks) {
      return false;
    }

    return this.nextLinkString in tasks;
  }

  public async loadMoreAndGetTasks(folderId: string) {
    let tasks = this.tasks[folderId];
    if (!tasks) {
      return this.getTasks(folderId);
    }

    if (this.nextLinkString in tasks) {
      const nextResource = tasks[this.nextLinkString].split('beta')[1];
      const response = await this.graph
        .api(nextResource)
        .version('beta')
        .middlewareOptions(prepScopes(this.scopes.read))
        .get();

      if (response && response.value) {
        response.value = tasks.value.concat(response.value);
        this.tasks[folderId] = response;
        tasks = response;
      }
    }

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof BaseGraph
   */
  public async completeTask(task: MicrosoftGraphBeta.OutlookTask): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.setTaskDetails(task.id, {
      isReminderOn: false,
      status: 'completed'
    });
  }

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof BaseGraph
   */
  public async incompleteTask(task: MicrosoftGraphBeta.OutlookTask): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.setTaskDetails(task.id, {
      isReminderOn: true,
      status: 'notStarted'
    });
  }

  public onUpdated(eventHandler: EventHandler<Event>) {
    this.taskUpdatedEventDispatcher.add(eventHandler);
  }

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public async removeTodoTask(task: MicrosoftGraphBeta.OutlookTask): Promise<any> {
    const response = await this.graph
      .api(`/me/outlook/tasks/${task.id}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.write))
      .delete();

    for (const folderId in this.tasks) {
      if (this.tasks.hasOwnProperty(folderId)) {
        const tasks = this.tasks[folderId].value;
        this.tasks[folderId].value = tasks.filter(t => t.id !== task.id);
      }
    }

    this.taskUpdatedEventDispatcher.fire(null);
  }

  /**
   * async promise, allows developer to add new to-do task
   *
   * @param {*} task
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof BaseGraph
   */
  public async addTask(task: MicrosoftGraphBeta.OutlookTask): Promise<MicrosoftGraphBeta.OutlookTask> {
    const { parentFolderId = null } = task;
    let newTask: MicrosoftGraphBeta.OutlookTask;

    if (parentFolderId) {
      newTask = await this.graph
        .api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
        .header('Cache-Control', 'no-store')
        .version('beta')
        .middlewareOptions(prepScopes(this.scopes.write))
        .post(task);
    } else {
      newTask = await this.graph
        .api('/me/outlook/tasks')
        .header('Cache-Control', 'no-store')
        .version('beta')
        .middlewareOptions(prepScopes(this.scopes.write))
        .post(task);
    }

    if (newTask) {
      if (this.tasks[newTask.parentFolderId]) {
        const tasks = this.tasks[newTask.parentFolderId];
        tasks.value = [newTask, ...tasks.value];
        this.taskUpdatedEventDispatcher.fire(null);
      }
    }

    return newTask;
  }

  private async getSingleTodoGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskGroup> {
    const group = await this.graph
      .api(`/me/outlook/taskGroups/${groupId}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return group;
  }

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTaskFolder[]>}
   * @memberof BaseGraph
   */
  private async getFoldersForTodoGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskFolder[]> {
    const folders = await this.graph
      .api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return folders && folders.value;
  }

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask[]>}
   * @memberof BaseGraph
   */
  private async getAllTodoTasksForFolder(folderId: string): Promise<MicrosoftGraphBeta.OutlookTask[]> {
    const tasks = await this.graph
      .api(`/me/outlook/taskFolders/${folderId}/tasks`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof BaseGraph
   */
  private async setTaskDetails(taskId: string, task: any): Promise<MicrosoftGraphBeta.OutlookTask> {
    const updatedTask = await this.graph
      .api(`/me/outlook/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes(this.scopes.write))
      .patch(task);

    for (const folderId in this.tasks) {
      if (this.tasks.hasOwnProperty(folderId)) {
        const tasks = this.tasks[folderId].value;
        const index = tasks.findIndex(t => t.id === updatedTask.id);
        if (index >= 0) {
          tasks[index] = updatedTask;
        }
      }
    }

    this.taskUpdatedEventDispatcher.fire(null);

    return updatedTask;
  }
}
