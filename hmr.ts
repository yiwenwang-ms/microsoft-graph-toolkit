import { customElement, html, LitElement, property } from 'lit-element';

@customElement('my-element')
export class MyElement extends LitElement {
  @property() public firstName;
  @property() public lastName;

  public render() {
    return html`
      <h1 style="color:yellow;">${this.firstName} ${this.lastName}</h1>
    `;
  }
}
