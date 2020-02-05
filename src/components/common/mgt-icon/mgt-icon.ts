import { customElement, html, LitElement, property } from 'lit-element';
import { styles } from './mgt-icon-css';

/**
 * Represents an Icon
 *
 * @export
 * @class mgt-icon
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-icon')
export class MgtIcon extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Fabric Core Icon Name
   * https://developer.microsoft.com/en-us/fabric#/styles/web/icons
   *
   * @type {string}
   * @memberof MgtIcon
   */
  @property({ attribute: 'name' }) public name: string;

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <i class="ms-Icon ms-Icon--${this.name}" aria-hidden="true"></i>
    `;
  }
}
