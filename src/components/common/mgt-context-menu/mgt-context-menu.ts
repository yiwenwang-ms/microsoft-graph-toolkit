/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import '../../common/mgt-button/mgt-button';
import { ButtonType } from '../../common/mgt-button/mgt-button';
import { MgtFlyout } from '../../common/mgt-flyout/mgt-flyout';
import { styles } from './mgt-context-menu-css';

/**
 * Context Menu Option
 *
 * @export
 * @interface ContextMenuOption
 */
export interface ContextMenuOption {
  /**
   * Unique option key
   *
   * @type {string}
   * @memberof ContextMenuOption
   */
  key: string;

  /**
   * Option text
   *
   * @type {string}
   * @memberof ContextMenuOption
   */
  text?: string;

  /**
   * Option icon
   *
   * @type {string}
   * @memberof ContextMenuOption
   */
  icon?: string;

  /**
   * Function called when option clicked
   *
   * @memberof ContextMenuOption
   */
  onClick: () => void;
}

/**
 * Custom Component used to handle an arrow rendering for TaskGroups utilized in the task component.
 *
 * @export MgtContextMenu
 * @class MgtContextMenu
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-context-menu')
export class MgtContextMenu extends MgtFlyout {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  public static get styles() {
    return [...MgtFlyout.styles, ...styles];
  }

  /**
   * Menu options to be rendered with an attached MouseEvent handler for expansion of details
   *
   * @memberof MgtContextMenu
   */
  @property({ type: Object }) public options: ContextMenuOption[];

  private focusedOptionIndex: number = 0;

  constructor() {
    super();

    this.addEventListener('keydown', this.handleFocus.bind(this));
  }

  protected updated(changedProps: PropertyValues) {
    super.updated(changedProps);

    if (this.isOpen) {
      this.focusedOptionIndex = 0;
      this.focusOption(this.focusedOptionIndex);
    }
  }

  // tslint:disable-next-line: completed-docs
  protected renderFlyout() {
    return html`
      <div class="options">
        ${repeat(this.options, o => o.key, o => this.renderOption(o))}
      </div>
    `;
  }

  /**
   * Renders an Context Menu Option
   *
   * @protected
   * @param {ContextMenuOption} option
   * @returns
   * @memberof MgtContextMenu
   */
  protected renderOption(option: ContextMenuOption) {
    return html`
      <mgt-button
        class="option"
        @click="${e => {
          if (option.onClick) {
            option.onClick();
          }
          this.isOpen = false;
        }}"
        .text=${option.text}
        .type=${ButtonType.command}
        icon="${option.icon}"
      >
      </mgt-button>
    `;
  }

  private focusOption(index: number) {
    const optionElems = this.renderRoot.querySelectorAll('.option');
    if (optionElems.length > index) {
      setTimeout(() => {
        (optionElems[index] as HTMLElement).focus();
      }, 20);
    }
  }

  private handleFocus(e: KeyboardEvent) {
    if (!this.isOpen) {
      return;
    }

    if (e.key === 'Tab' || e.keyCode === 9) {
      if (e.shiftKey) {
        this.focusedOptionIndex = (this.focusedOptionIndex - 1 + this.options.length) % this.options.length;
      } else {
        this.focusedOptionIndex = (this.focusedOptionIndex + 1) % this.options.length;
      }
      this.focusOption(this.focusedOptionIndex);
      e.preventDefault();
    } else if (e.keyCode === 40) {
      // down arrow
      this.focusedOptionIndex = (this.focusedOptionIndex + 1) % this.options.length;
      this.focusOption(this.focusedOptionIndex);
      e.preventDefault();
    } else if (e.keyCode === 38) {
      // up arrow
      this.focusedOptionIndex = (this.focusedOptionIndex - 1 + this.options.length) % this.options.length;
      this.focusOption(this.focusedOptionIndex);
      e.preventDefault();
    }
  }
}
