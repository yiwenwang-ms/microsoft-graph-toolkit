import { css, customElement, html, LitElement, property } from 'lit-element';
import { styleMap } from 'lit-html/directives/style-map';
import { styles } from './mgt-shimmer-css';

/**
 *
 *
 * @export
 * @class mgt-shimmer
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-shimmer')
export class MgtShimmer extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  // assignment to this property will re-render the component
  @property() private width: string;
  @property() private height: string;

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const children = this.children;
    const shimmerElements = [];

    let rowHeight = 0;

    if (!children.length) {
      shimmerElements.push(this.renderLine());
    } else {
      for (let i = 0; i < children.length; i++) {
        const child = children[i] as MgtShimmerElement;
        let height = 0;

        switch (child.tagName) {
          case 'MGT-SHIMMER-LINE':
            height = child.height || MgtShimmerLine.defaultHeight;
            break;
          case 'MGT-SHIMMER-GAP':
            height = child.height || MgtShimmerGap.defaultHeight;
            break;
          case 'MGT-SHIMMER-CIRCLE':
            height = child.height || MgtShimmerCircle.defaultHeight;
            break;
        }
        rowHeight = Math.max(rowHeight, height);
      }
      for (let i = 0; i < children.length; i++) {
        const child = children[i] as MgtShimmerElement;

        switch (child.tagName) {
          case 'MGT-SHIMMER-LINE':
            shimmerElements.push(
              this.renderLine({
                height: child.height || MgtShimmerLine.defaultHeight,
                width: child.width,
                style: this.getStyles(child.verticalAlign, child.height || MgtShimmerLine.defaultHeight, rowHeight)
              })
            );
            break;
          case 'MGT-SHIMMER-GAP':
            shimmerElements.push(this.renderGap(child.width, rowHeight));
            break;
          case 'MGT-SHIMMER-CIRCLE':
            shimmerElements.push(
              this.renderCircle({
                height: child.height || MgtShimmerCircle.defaultHeight,
                width: child.width,
                style: this.getStyles(child.verticalAlign, child.height || MgtShimmerCircle.defaultHeight, rowHeight)
              })
            );
            break;
        }
      }
    }

    const wrapperStyle = {
      width: this.width
    };

    return html`
      <div class="container">
        <div class="wrapper" style=${styleMap(wrapperStyle)}>
          <div class="gradient"></div>
          <div class="root">
            ${shimmerElements}
          </div>
        </div>
      </div>
    `;
  }

  protected renderLine(options?: { style: any; width?: number | string; height?: number }) {
    let { style, width, height } = options || {};

    style = style || {};

    if (width) {
      style.width = typeof width === 'number' ? width + 'px' : width;
      style['min-width'] = style.width;
    }

    if (height) {
      style.height = height + 'px';
    }

    width = width || height;

    return html`
      <div class="line" style=${styleMap(style)}></div>
    `;
  }

  protected renderGap(width?: string | number, height?: number) {
    const style: any = {};

    style.width = width ? (typeof width === 'number' ? width + 'px' : width) : null;
    style.height = height ? height + 'px' : null;

    return html`
      <div class="gap" style=${styleMap(style)}></div>
    `;
  }

  protected renderCircle(options: { style: any; width?: number | string; height?: number }) {
    let { style, width, height } = options;

    style = style || {};

    if (width) {
      style.width = typeof width === 'number' ? width + 'px' : width;
      style['min-width'] = style.width;
    }

    if (height) {
      style.height = height + 'px';
    }

    width = width || height;

    return html`
      <div class="circle" style=${styleMap(style)}>
        <svg viewBox="0 0 10 10" width=${width} height=${height}>
          <path
            d="M0,0 L10,0 L10,10 L0,10 L0,0 Z M0,5 C0,7.76142375 2.23857625,10 5,10 C7.76142375,10 10,7.76142375 10,5 C10,2.23857625 7.76142375,2.22044605e-16 5,0 C2.23857625,-2.22044605e-16 0,2.23857625 0,5 L0,5 Z"
          />
        </svg>
      </div>
    `;
  }

  private getStyles(verticalAlign: ShimmerVerticalAlign, elementHeight: number, rowHeight: number) {
    const dif: number = rowHeight && elementHeight ? rowHeight - elementHeight : 0;

    let style: any;

    if (verticalAlign === ShimmerVerticalAlign.center) {
      style = {
        borderBottomWidth: `${dif ? Math.floor(dif / 2) : 0}px`,
        borderTopWidth: `${dif ? Math.ceil(dif / 2) : 0}px`
      };
    } else if (verticalAlign === ShimmerVerticalAlign.top) {
      style = {
        borderBottomWidth: `${dif}px`,
        borderTopWidth: '0px'
      };
    } else if (verticalAlign === ShimmerVerticalAlign.bottom) {
      style = {
        borderBottomWidth: '0px',
        borderTopWidth: `${dif}px`
      };
    }

    return style;
  }
}

export enum ShimmerVerticalAlign {
  top,
  center,
  bottom
}

class MgtShimmerElement extends LitElement {
  @property()
  public width: number | string;

  @property()
  public height: number;

  @property({
    attribute: 'vertical-align',
    converter: (value, type) => {
      value = value.toLowerCase();
      if (typeof ShimmerVerticalAlign[value] === 'undefined') {
        return ShimmerVerticalAlign.center;
      } else {
        return ShimmerVerticalAlign[value];
      }
    }
  })
  public verticalAlign: ShimmerVerticalAlign;

  public createRenderRoot() {
    return null;
  }
}

@customElement('mgt-shimmer-line')
export class MgtShimmerLine extends MgtShimmerElement {
  public static readonly defaultHeight = 16;
}

@customElement('mgt-shimmer-gap')
export class MgtShimmerGap extends MgtShimmerElement {
  public static readonly defaultHeight = 16;
}

@customElement('mgt-shimmer-circle')
export class MgtShimmerCircle extends MgtShimmerElement {
  public static readonly defaultHeight = 50;
}
