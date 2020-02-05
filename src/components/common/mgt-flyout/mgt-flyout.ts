/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, LitElement, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { styles } from './mgt-flyout-css';

/**
 * A component to create flyout anchored to an element
 *
 * @export
 * @class MgtFlyout
 * @extends {LitElement}
 */
@customElement('mgt-flyout')
export class MgtFlyout extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Gets or sets whether the flyout is visible
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({ attribute: 'isOpen', type: Boolean })
  get isOpen() {
    return this.privateIsOpen;
  }

  set isOpen(value: boolean) {
    const oldValue = this.privateIsOpen;
    this.privateIsOpen = value;

    setTimeout(() => {
      if (value) {
        window.addEventListener('wheel', this.handleWindowEvent);
        window.addEventListener('pointerdown', this.handleWindowEvent);
        window.addEventListener('resize', this.handleResize);
      } else {
        window.removeEventListener('wheel', this.handleWindowEvent);
        window.removeEventListener('pointerdown', this.handleWindowEvent);
        window.removeEventListener('resize', this.handleResize);
      }
    }, 20);

    this.requestUpdate('isOpen', oldValue);
  }

  private privateIsOpen: boolean = false;
  private renderedOnce: boolean = false;

  public constructor() {
    super();

    this.handleWindowEvent = this.handleWindowEvent.bind(this);
    this.handleResize = this.handleResize.bind(this);

    this.addEventListener('keyup', this.handleKeyUp);
  }

  /**
   * Invoked whenever the element is updated. Implement to perform
   * post-updating tasks via DOM APIs, for example, focusing an element.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param changedProperties Map of changed properties with old values
   */
  protected updated(changedProps: PropertyValues) {
    super.updated(changedProps);

    window.requestAnimationFrame(() => {
      this.updateFlyout();
    });
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const classes = {
      root: true,
      visible: this.isOpen
    };

    let flyout = null;

    if (this.isOpen || this.renderedOnce) {
      this.renderedOnce = true;
      flyout = html`
        <div class="flyout">
          ${this.renderFlyout()}
        </div>
      `;
    }

    return html`
      <div class=${classMap(classes)}>
        <div class="anchor">
          ${this.renderAnchor()}
        </div>
        ${flyout}
      </div>
    `;
  }

  /**
   * Renders the anchor content
   *
   * @protected
   * @returns
   * @memberof MgtFlyout
   */
  protected renderAnchor() {
    return html`
      <slot></slot>
    `;
  }

  /**
   * Renders flyout and smoke
   *
   * @protected
   * @returns
   * @memberof MgtFlyout
   */
  protected renderFlyout() {
    return html`
      <slot name="flyout"></slot>
    `;
  }

  private handleWindowEvent(e: Event) {
    const flyout = this.renderRoot.querySelector('.flyout');

    if (flyout) {
      // IE
      if (!e.composedPath) {
        let currentElem = e.target as HTMLElement;
        while (currentElem) {
          currentElem = currentElem.parentElement;
          if (currentElem === flyout || (e.type === 'pointerdown' && currentElem === this)) {
            return;
          }
        }
      } else {
        const path = e.composedPath();
        if (path.includes(flyout) || (e.type === 'pointerdown' && path.includes(this))) {
          return;
        }
      }
    }

    this.close();
  }

  private handleResize(e: Event) {
    this.close();
  }

  private handleKeyUp(e: KeyboardEvent) {
    if (e.key === 'Escape') {
      this.close();
    }
  }

  private close() {
    this.isOpen = false;
    this.dispatchEvent(new Event('closed'));
  }

  private updateFlyout() {
    if (!this.isOpen) {
      return;
    }

    const anchor = this.renderRoot.querySelector('.anchor');
    const flyout = this.renderRoot.querySelector('.flyout') as HTMLElement;

    const windowWidth =
      window.innerWidth && document.documentElement.clientWidth
        ? Math.min(window.innerWidth, document.documentElement.clientWidth)
        : window.innerWidth || document.documentElement.clientWidth;

    const windowHeight =
      window.innerHeight && document.documentElement.clientHeight
        ? Math.min(window.innerHeight, document.documentElement.clientHeight)
        : window.innerHeight || document.documentElement.clientHeight;

    if (flyout && anchor) {
      let left: number;
      // let bottom: number;
      let top: number;

      const flyoutRect = flyout.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();

      // normalize flyoutrect since we could have moved it before
      // need to know where would it render, not where it renders
      // const flyoutTop = anchorRect.bottom;
      const flyoutLeft = anchorRect.left;
      const flyoutRight = flyoutLeft + flyoutRect.width;
      // const flyoutBottom = flyoutTop + flyoutRect.height;

      // if (flyoutRect.width > windowWidth) {
      //   // page width is smaller than flyout, render all the way to the left
      //   left = -flyoutLeft;
      // } else if (anchorRect.width >= flyoutRect.width) {
      //   // anchor is large than flyout, render aligned to anchor
      //   left = 0;
      // } else {
      //   const centerOffset = flyoutRect.width / 2 - anchorRect.width / 2;

      //   if (flyoutLeft - centerOffset < 20) {
      //     // centered flyout is off screen to the left, render on the left edge
      //     left = -flyoutLeft + 20;
      //   } else if (flyoutRight - centerOffset > windowWidth - 20) {
      //     // centered flyout is off screen to the right, render on the right edge
      //     left = -(flyoutRight - windowWidth) - 20;
      //   } else {
      //     // render centered
      //     left = -centerOffset;
      //   }
      // }

      // if (flyoutRect.height > windowHeight || (windowHeight < flyoutBottom && anchorRect.top < flyoutRect.height)) {
      //   top = -flyoutTop + anchorRect.height;
      // } else if (windowHeight < flyoutBottom) {
      //   // bottom = anchorRect.height;
      // }

      // TODO (when flyout is off screen when rendered above)
      if (flyoutRect.height > windowHeight) {
        top = 0;
      } else if (anchorRect.y + anchorRect.height + flyoutRect.height >= windowHeight) {
        // it will be off screen, render above
        top = anchorRect.y - flyoutRect.height;
      } else {
        top = anchorRect.y + anchorRect.height;
      }

      if (flyoutRect.width > windowWidth) {
        left = 0;
      } else if (anchorRect.x + flyoutRect.width >= windowWidth) {
        left = anchorRect.x - (flyoutRect.width - anchorRect.width);
      } else {
        left = anchorRect.x;
      }

      flyout.style.left = typeof left !== 'undefined' ? `${left}px` : '';
      flyout.style.top = typeof top !== 'undefined' ? `${top}px` : '';
    }
  }
}
