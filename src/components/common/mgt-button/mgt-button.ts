import { customElement, html, LitElement, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import '../mgt-icon/mgt-icon';
import { styles } from './mgt-button-css';

/**
 * Defines how a person card is shown when a user interacts with
 * a person component
 *
 * @export
 * @enum {number}
 */
export enum ButtonType {
  /**
   * Default button
   */
  default,

  /**
   * Primary button
   */
  primary,

  /**
   * Command button
   */
  command
}

/**
 * A button
 *
 * @export
 * @class mgt-button
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-button')
export class MgtButton extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * The text to display
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({ attribute: 'text' }) public text: string;

  /**
   * Fabric Core Icon Name
   * https://developer.microsoft.com/en-us/fabric#/styles/web/icons
   *
   * @type {string}
   * @memberof MgtButton
   */
  @property({ attribute: 'icon' }) public icon: string;

  /**
   * The type of button
   *
   * @type {string}
   * @memberof MgtButton
   */
  @property({
    attribute: 'type',
    converter: (value, type) => {
      value = value.toLowerCase();
      return ButtonType[value] || ButtonType.default;
    }
  })
  public type: ButtonType;

  /**
   * Sets whether the button is disabled
   *
   * @type {boolean}
   * @memberof MgtButton
   */
  @property({ attribute: 'disabled' }) public disabled: boolean;

  constructor() {
    super();

    this.type = ButtonType.default;
  }

  /** @inheritdoc */
  // tslint:disable-next-line: completed-docs
  public focus(focusOptions?: FocusOptions) {
    const button = this.renderRoot.querySelector('button');
    if (button) {
      button.focus(focusOptions);
    }
  }

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
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const className = {
      disabled: this.disabled,
      root: true
    };

    if (!this.disabled) {
      className[ButtonType[this.type]] = true;
    }

    let content;

    if (this.children.length) {
      content = html`
        <slot></slot>
      `;
    } else {
      content = html`
        ${this.renderIcon()} ${this.renderText()}
      `;
      // tslint:disable-next-line: no-string-literal
      className['no-content'] = true;
    }

    return html`
      <button class=${classMap(className)} .disabled=${this.disabled}>
        <div class="content">
          ${content}
        </div>
      </button>
    `;
  }

  /**
   * Returns the icon element if defined
   *
   * @protected
   * @returns
   * @memberof MgtButton
   */
  protected renderIcon() {
    if (this.icon) {
      const className = this.text && this.text !== '' ? 'with-text' : '';
      return html`
        <mgt-icon class="${className}" name="${this.icon}"></mgt-icon>
      `;
    }

    return null;
  }

  /**
   * Returns the text element if defined
   *
   * @protected
   * @returns
   * @memberof MgtButton
   */
  protected renderText() {
    if (this.text && this.text !== '') {
      return html`
        <span class="text">${this.text}</span>
      `;
    }
  }
}
