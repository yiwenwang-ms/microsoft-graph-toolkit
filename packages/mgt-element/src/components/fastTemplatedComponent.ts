/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { equals } from '../utils/equals';
import { TemplateContext } from '../utils/TemplateContext';
import { TemplateHelper } from '../utils/TemplateHelper';
import { MgtFastBaseComponent } from './fastBaseComponent';
import { html, ViewTemplate } from '@microsoft/fast-element';

/**
 * Lookup for rendered component templates and contexts by slot name.
 */
interface RenderedTemplates {
  [name: string]: {
    /**
     * Reference to the data context used to render the slot.
     */
    context: any;
    /**
     * Reference to the rendered DOM element corresponding to the slot.
     */
    slot: HTMLElement;
  };
}

/**
 * An abstract class that defines a templatable web component
 *
 * @export
 * @abstract
 * @class MgtTemplatedComponent
 * @extends {MgtBaseComponent}
 */
export abstract class MgtFastTemplatedComponent extends MgtFastBaseComponent {
  /**
   * Additional data context to be used in template binding
   * Use this to add event listeners or value converters
   *
   * @type {MgtElement.TemplateContext}
   * @memberof MgtTemplatedComponent
   */
  public templateContext: TemplateContext;

  /**
   * Holds all templates defined by developer
   *
   * @protected
   * @memberof MgtTemplatedComponent
   */
  protected templates = {};

  private _renderedSlots = false;
  private _renderedTemplates: RenderedTemplates = {};
  private _slotNamesAddedDuringRender = [];

  constructor() {
    super();

    this.templateContext = this.templateContext || {};
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtBaseComponent
   */
  public connectedCallback() {
    this.templates = this.getTemplates();
    this._slotNamesAddedDuringRender = [];
    super.connectedCallback();

    // TODO: Could this be a way to clean up slots???
    this.$fastController.subscribe({ handleChange: () => console.log('hello') }, 'events');
  }

  // TODO
  // /**
  //  * Invoked whenever the element is updated. Implement to perform
  //  * post-updating tasks via DOM APIs, for example, focusing an element.
  //  *
  //  * Setting properties inside this method will trigger the element to update
  //  * again after this update cycle completes.
  //  *
  //  * * @param changedProperties Map of changed properties with old values
  //  */
  // protected updated(changedProperties: PropertyValues) {
  //   super.updated(changedProperties);
  //   this.removeUnusedSlottedElements();
  // }

  /**
   * Render a <template> by type and return content to render
   *
   * @param templateType type of template (indicated by the data-type attribute)
   * @param context the data context that should be expanded in template
   * @param slotName the slot name that will be used to host the new rendered template. set to a unique value if multiple templates of this type will be rendered. default is templateType
   */
  public getTemplate(templateType: string, context: object, slotName: string, defaultTemplate: ViewTemplate) {
    if (!this.hasTemplate(templateType)) {
      return defaultTemplate;
    }

    slotName = slotName || templateType;
    this._slotNamesAddedDuringRender.push(slotName);
    this._renderedSlots = true;

    const template = html`
      <slot name=${slotName}></slot>
    `;

    if (this._renderedTemplates.hasOwnProperty(slotName)) {
      const { context: existingContext, slot } = this._renderedTemplates[slotName];
      if (equals(existingContext, context)) {
        return template;
      }
      this.removeChild(slot);
    }

    const div = document.createElement('div');
    div.slot = slotName;
    div.dataset.generated = 'template';

    TemplateHelper.renderTemplate(div, this.templates[templateType], context, {
      ...this.templateContext
    });

    this.appendChild(div);

    this._renderedTemplates[slotName] = { context, slot: div };

    this.$emit('templateRendered', { templateType, context, element: div });

    return template;
  }

  /**
   * Check if a specific template has been provided.
   *
   * @protected
   * @param {string} templateName
   * @returns {boolean}
   * @memberof MgtTemplatedComponent
   */
  protected hasTemplate(templateName: string): boolean {
    return this.templates && this.templates[templateName];
  }

  private getTemplates() {
    const templates: any = {};

    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < this.children.length; i++) {
      const child = this.children[i];
      if (child.nodeName === 'TEMPLATE') {
        const template = child as HTMLElement;
        if (template.dataset.type) {
          templates[template.dataset.type] = template;
        } else {
          templates.default = template;
        }

        (template as any).templateOrder = i;
      }
    }

    return templates;
  }

  // TODO - how do we know when to remove children elements from element
  private removeUnusedSlottedElements() {
    if (this._renderedSlots) {
      for (let i = 0; i < this.children.length; i++) {
        const child = this.children[i] as HTMLElement;
        if (child.dataset && child.dataset.generated && !this._slotNamesAddedDuringRender.includes(child.slot)) {
          this.removeChild(child);
          delete this._renderedTemplates[child.slot];
          i--;
        }
      }
      this._renderedSlots = false;
    }
  }
}
