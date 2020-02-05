import { GraphRequest } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { Graph } from '../../../Graph';
import { EventDispatcher, EventHandler } from '../../../utils/EventDispatcher';
import { prepScopes } from '../../../utils/GraphHelpers';

/**
 * TODO
 *
 * @export
 * @class ToDoTaskSource
 */
export interface BaseTaskSource<T, F, G> {
  onUpdated(eventHandler: EventHandler<Event>);

  getFolders(): Promise<F[]>;

  getTasks(folderId: string): Promise<T[]>;
  hasMoreForFolder(folderId: string): boolean;
  loadMoreAndGetTasks(folderId: string): Promise<T[]>;

  addTask(task: T): Promise<T>;
}
