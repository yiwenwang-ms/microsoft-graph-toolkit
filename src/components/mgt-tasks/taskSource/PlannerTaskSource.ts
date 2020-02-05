import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Graph } from '../../../Graph';
import { EventDispatcher, EventHandler } from '../../../utils/EventDispatcher';
import { prepScopes } from '../../../utils/GraphHelpers';

/**
 * TODO
 *
 * @export
 * @class ToDoTaskSource
 */
export class PlannerTaskSource {
  private readonly scopes = {
    read: 'Group.Read.All',
    write: 'Group.ReadWrite.All'
  };

  private readonly nextLinkString = '@odata.nextLink';

  private graph: Graph;
  private tasks: Record<string, { value: MicrosoftGraph.PlannerTask[] }> = {};

  private buckets: Record<string, any>;
  private plans: MicrosoftGraph.PlannerPlan[];

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
  public async getTasks(folderId): Promise<MicrosoftGraph.PlannerTask[]> {
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

  public async getPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    if (this.plans.length > 0) {
      return this.plans;
    }

    const plans = await this.graph
      .api('/me/planner/plans')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    this.plans = plans && plans.value;

    return this.plans;
  }

  // public async getFolders(): Promise<MicrosoftGraph.PlannerBucket[]> {
  //   if (this.folders.length > 0) {
  //     return this.folders;
  //   }

  //   const request = await this.graph
  //     .api('me/outlook/taskFolders')
  //     .version('beta')
  //     .middlewareOptions(prepScopes(this.scopes.read))
  //     .get();

  //   this.folders = request && request.value;
  //   return this.folders;
  // }

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

  public onUpdated(eventHandler: EventHandler<Event>) {
    this.taskUpdatedEventDispatcher.add(eventHandler);
  }

  private async getMyTasks(): Promise<MicrosoftGraph.PlannerTask[]> {
    const tasks = await this.graph
      .api('me/planner/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraph.PlannerPlan[]>}
   * @memberof BaseGraph
   */
  private async getPlansForGroup(groupId: string): Promise<MicrosoftGraph.PlannerPlan[]> {
    const uri = `/groups/${groupId}/planner/plans`;
    const plans = await this.graph
      .api(uri)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();
    return plans ? plans.value : null;
  }

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {Promise<MicrosoftGraph.PlannerPlan>}
   * @memberof BaseGraph
   */
  private async getSinglePlannerPlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const plan = await this.graph
      .api(`/planner/plans/${planId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return plan;
  }

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {Promise<MicrosoftGraph.PlannerBucket[]>}
   * @memberof BaseGraph
   */
  private async getBucketsForPlannerPlan(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    const buckets = await this.graph
      .api(`/planner/plans/${planId}/buckets`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return buckets && buckets.value;
  }

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {Promise<MicrosoftGraph.PlannerTask[]>}
   * @memberof BaseGraph
   */
  private async getTasksForPlannerBucket(bucketId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    const tasks = await this.graph
      .api(`/planner/buckets/${bucketId}/tasks`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.read))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to set details of planner task associated with a taskId
   *
   * @param {string} taskId
   * @param {MicrosoftGraph.PlannerTask} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async setPlannerTaskDetails(taskId: string, details: MicrosoftGraph.PlannerTask, eTag: string): Promise<any> {
    return await this.graph
      .api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.write))
      .header('If-Match', eTag)
      .patch(JSON.stringify(details));
  }

  /**
   * async promise, allows developer to set a task to complete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async setPlannerTaskComplete(taskId: string, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
      taskId,
      {
        percentComplete: 100
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to set a task to incomplete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async setPlannerTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
      taskId,
      {
        percentComplete: 0
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to assign people to task
   *
   * @param {string} taskId
   * @param {*} people
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async assignPeopleToPlannerTask(taskId: string, people: any, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
      taskId,
      {
        assignments: people
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to create new Planner task
   *
   * @param {MicrosoftGraph.PlannerTask} newTask
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async addPlannerTask(newTask: MicrosoftGraph.PlannerTask): Promise<any> {
    return this.graph
      .api('/planner/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(this.scopes.write))
      .post(newTask);
  }

  /**
   * async promise, allows developer to remove Planner task associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  private async removePlannerTask(taskId: string, eTag: string): Promise<any> {
    return this.graph
      .api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes(this.scopes.write))
      .delete();
  }
}
