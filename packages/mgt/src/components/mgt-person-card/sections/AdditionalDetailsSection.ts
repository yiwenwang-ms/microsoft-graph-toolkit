import { BasePersonCardSection } from './BasePersonCardSection';
import { TemplateResult, customElement } from 'lit-element';

@customElement('mgt-person-card-additional-details')
export class AdditionalDetailsSection extends BasePersonCardSection {
  /**
   * The name of the section
   *
   * @readonly
   * @type {string}
   * @memberof AdditionalDetailsSection
   */
  public get displayName(): string {
    return 'Additional Details';
  }

  /**
   * Render the section icon
   *
   * @memberof AdditionalDetailsSection
   */
  public renderIcon: () => TemplateResult;

  /**
   * Clear any state from the section
   *
   * @memberof AdditionalDetailsSection
   */
  public clearState(): void {}

  /**
   * Render the section in compact mode
   *
   * @memberof AdditionalDetailsSection
   */
  public renderCompactView: () => TemplateResult;

  /**
   * Render the section in full mode
   *
   * @memberof AdditionalDetailsSection
   */
  public renderFullView: () => TemplateResult;

  constructor(
    renderIcon: () => TemplateResult,
    renderCompactView: () => TemplateResult,
    renderFullView: () => TemplateResult
  ) {
    super();
    this.renderIcon = renderIcon;
    this.renderCompactView = renderCompactView;
    this.renderFullView = renderFullView;
  }
}
