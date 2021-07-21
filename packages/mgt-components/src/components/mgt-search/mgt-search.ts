/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import '../../styles/style-helper';
import '../sub-components/mgt-spinner/mgt-spinner';
import { debounce } from '../../utils/Utils';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { styles } from './mgt-search-css';

import { strings } from './strings';
import {
  getSuggestions,
  SuggestionFile,
  SuggestionPeople,
  Suggestions,
  SuggestionQuery,
  SuggestionConfig,
  SuggestionEntityConfig
} from '../../graph/graph.search';
import { IDynamicPerson } from '../../graph/types';

/**
 * An interface used to mark an object as 'focused',
 * so it can be rendered differently.
 *
 * @interface IFocusable
 */
interface IFocusable {
  // tslint:disable-next-line: completed-docs
  isFocused: boolean;
}

/**
 * Web component used to search for people from the Microsoft Graph
 *
 * @export
 * @fires suggestionClick - Fired when selection changes
 * @fires enterPress - Fired when selection changes
 * @cssprop --suggestion-item-background-color--hover - {Color} background color for an hover item
 * @cssprop --suggestion-list-background-color - {Color} background color
 * @cssprop --suggestion-list-query-color - {Color} Query Suggestion font color
 *
 */
@customElement('mgt-search')
export class MgtSearch extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * Gets the input element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get input(): HTMLInputElement {
    return this.renderRoot.querySelector('.search-box__input');
  }

  /**
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({ attribute: 'value' })
  public get value() {
    return this.input.value;
  }

  public set value(_value) {
    this.input.value = _value;
  }

  @property({ attribute: 'suggestion-value' })
  public get suggestionValue() {
    return this.getFocusItemValue();
  }

  /**
   * Placeholder query.
   *
   * @type {string}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled: boolean;

  @property({
    attribute: 'people-suggestion-value',
    type: String
  })
  private _peopleSuggestionValue: string;

  public set peopleSuggestionValue(peopleSuggestionValue: string) {
    this._peopleSuggestionValue = peopleSuggestionValue;
  }

  public get peopleSuggestionValue() {
    return this._peopleSuggestionValue;
  }

  @property({
    attribute: 'file-suggestion-value',
    type: String
  })
  private _fileSuggestionValue: string;

  public set fileSuggestionValue(fileSuggestionValue: string) {
    this._fileSuggestionValue = fileSuggestionValue;
  }

  public get fileSuggestionValue() {
    return this._fileSuggestionValue;
  }

  @property({
    attribute: 'query-suggestion-value',
    type: String
  })
  private _querySuggestionValue: string;

  public set querySuggestionValue(querySuggestionValue: string) {
    this._querySuggestionValue = querySuggestionValue;
  }

  public get querySuggestionValue() {
    return this._querySuggestionValue;
  }

  @property({
    attribute: 'max-file-suggestions',
    type: Number
  })
  private _maxFileSuggestionCount: number = 3;

  public set maxFileSuggestionCount(maxFileSuggestionCount: number) {
    this._maxFileSuggestionCount = maxFileSuggestionCount;
  }

  public get maxFileSuggestionCount() {
    return this._maxFileSuggestionCount;
  }

  @property({
    attribute: 'max-query-suggestions',
    type: Number
  })
  private _maxQuerySuggestionCount: number = 3;

  public set maxQuerySuggestionCount(maxQuerySuggestionCount: number) {
    this._maxQuerySuggestionCount = maxQuerySuggestionCount;
  }

  public get maxQuerySuggestionCount() {
    return this._maxQuerySuggestionCount;
  }

  @property({
    attribute: 'max-people-suggestions',
    type: Number
  })
  private _maxPeopleSuggestionCount: number = 3;

  public set maxPeopleSuggestionCount(maxPeopleSuggestionCount: number) {
    this._maxPeopleSuggestionCount = maxPeopleSuggestionCount;
  }

  public get maxPeopleSuggestionCount() {
    return this._maxPeopleSuggestionCount;
  }

  @property({
    attribute: 'entity-types',
    type: String
  })
  private _entityTypes: string = 'file, query, people';

  public set entityTypes(entityTypes: string) {
    this._entityTypes = entityTypes;
  }

  @property({
    attribute: 'cvid',
    type: String
  })
  private _cvid: string = 'd8e48cff-9cac-40c9-0b5c-e6d24488f781';

  public set cvid(cvid: string) {
    this._cvid = cvid;
  }

  @property({
    attribute: 'text-decorations',
    type: String
  })
  private _textDecorations: string = 'd8e48cff-9cac-40c9-0b5c-e6d24488f781';

  private _suggestionConfig: SuggestionConfig;

  private _suggestionEntityOrder = [];

  public set textDecorations(textDecorations: string) {
    this._textDecorations = textDecorations;
  }

  public get entityTypes() {
    return this._entityTypes;
  }

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtSearch
   */
  protected userInput: string;

  // if search is still loading don't load "people not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  // tracking of user arrow key input for selection
  private _arrowSelectionCount: number = 0;

  private _debouncedSearch: { (): void; (): void };

  private arrow_initial = false;

  @internalProperty() private _isFocused = false;

  @internalProperty() private _foundPeopleSuggestion: SuggestionPeople[];

  @internalProperty() private _foundQuerySuggestion: SuggestionQuery[];

  @internalProperty() private _foundFileSuggestion: SuggestionFile[];

  @internalProperty() private _foundSuggestion: Map<string, any[]>;

  @internalProperty() private _suggestionItemsMap: Map<string, any> = new Map();

  constructor() {
    super();
    this.clearState();
    this._showLoading = true;
    this.disabled = false;
  }

  /**
   * Get the scopes required for search suggestion
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtSearch
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set(['']) // awaiting on search suggestion onboard
    ];
  }

  /**
   * Focuses the input element when focus is called
   *
   * @param {FocusOptions} [options]
   * @memberof MgtSearch
   */
  public focus(options?: FocusOptions) {
    this.gainedFocus();
    if (!this.input) {
      return;
    }
    this.input.focus(options);
    this.input.select();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns {TemplateResult}
   * @memberof MgtSearch
   */
  public render(): TemplateResult {
    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);

    const inputClasses = {
      focused: this._isFocused,
      search: true,
      disabled: this.disabled
    };

    return html`
      <div dir=${this.direction} class=${classMap(inputClasses)} @click=${e => this.focus(e)}>
        <div class="selected-list">
          ${flyoutTemplate}
        </div>
      </div>
    `;
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtSearch
   */
  protected clearState(): void {
    this.userInput = '';
  }

  /**
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected requestStateUpdate(force?: boolean) {
    return super.requestStateUpdate(force);
  }

  /**
   * Render the input query box.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearch
   */
  protected renderInput(): TemplateResult {
    const inputClasses = {
      'search-box': true
    };

    return (
      this.renderTemplate('suggestion-input', null) ||
      html`
      <div class="${classMap(inputClasses)}">
        <input
          id="suggestion-input"
          class="search-box__input"
          type="text"
          placeholder="Search sth..."
          label="suggestion-input"
          aria-label="suggestion-input"
          role="input"
          @keydown="${this.onUserKeyDown}"
          @keyup="${this.onUserKeyUp}"
          @blur=${this.lostFocus}
          @click=${this.handleFlyout}
          ?disabled=${this.disabled}
        />
      </div>
    `
    );
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtSearch
   */
  protected renderLoading(): TemplateResult {
    return (
      this.renderTemplate('loading', null) ||
      html`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-query" aria-label="loading" class="loading-query">
            ${this.strings.loadingMessage}
          </div>
        </div>
      `
    );
  }

  /**
   * Render the state when no results are found for the search query.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearch
   */
  protected renderNoData(): TemplateResult {
    return (
      this.renderTemplate('error', null) ||
      this.renderTemplate('no-data', null) ||
      html`
        <div class="message-parent">
          <div label="search-error-query" aria-label="We didn't find any matches." class="search-error-query">
            ${this.strings.noResultsFound}
          </div>
        </div>
      `
    );
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    const input = this.userInput;
    if (provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);
      if (this._isFocused) {
        this.setEntityConfig();
        var suggestions = await getSuggestions(graph, this._suggestionConfig);
        this._foundSuggestion = suggestions;
      }
      this._showLoading = false;
      this.clearArrowSelection();
    }
  }

  /**
   * Hide the results flyout.
   *
   * @protected
   * @memberof MgtSearch
   */
  protected hideFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtSearch
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }

  private clearInput() {
    this.input.value = '';
    this.userInput = '';
  }

  private handleFlyout() {
    // handles hiding control if default people have no more selections available
    let shouldShow = true;
    if (shouldShow) {
      window.requestAnimationFrame(() => {
        // Mouse is focused on input
        this.showFlyout();
      });
    }
  }

  private gainedFocus() {
    this._isFocused = true;
    if (this.input) {
      this.input.focus();
    }
    this._showLoading = true;
    this.loadState();
  }

  private lostFocus() {
    this._isFocused = false;
    this.requestUpdate();
  }

  /**
   * Adds debounce method for set delay on user input
   */
  private onUserKeyUp(event: KeyboardEvent): void {
    if (event.keyCode === 40 || event.keyCode === 39 || event.keyCode === 38 || event.keyCode === 37) {
      // keyCodes capture: down arrow (40), right arrow (39), up arrow (38) and left arrow (37)
      return;
    }

    const input = event.target as HTMLInputElement;

    if (event.code === 'Escape') {
      this.clearInput();
    } else {
      this.userInput = input.value;
      this.handleUserSearch();
    }
  }

  /**
   * Tracks event on user input in search
   * @param input - input query
   */
  private handleUserSearch() {
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        const loadingTimeout = setTimeout(() => {
          this._showLoading = true;
        }, 50);

        await this.loadState();
        clearTimeout(loadingTimeout);
        this._showLoading = false;
        this.showFlyout();

        this._arrowSelectionCount = 0;
      }, 400);
    }

    this._debouncedSearch();
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: KeyboardEvent): void {
    if (!this.flyout.isOpen) {
      return;
    }
    if (event.keyCode === 40 || event.keyCode === 38 || event.keyCode === 9) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this.input.value.length > 0) {
        event.preventDefault();
      }
    }
    if (event.keyCode === 13) {
      //  and enter (13)
      //this.onEnterKeyPressCallback(this.input.value, this.getFocusItemValue());
      this.fireCustomEvent('onEnterPress', {
        originalValue: this.input.value,
        suggestedValue: this.getFocusItemValue()
      });

      this.hideFlyout();
      (event.target as HTMLInputElement).value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: KeyboardEvent): void {
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    if (peopleList && peopleList.length) {
      // update arrow count
      if (event.keyCode === 38) {
        // up arrow
        if (this.arrow_initial) {
          this._arrowSelectionCount = (this._arrowSelectionCount - 1 + peopleList.length) % peopleList.length;
        } else {
          this.arrow_initial = true;
          this._arrowSelectionCount = peopleList.length - 1;
          this._arrowSelectionCount = 0;
        }
      }
      if (event.keyCode === 40 || event.keyCode === 9) {
        // down arrow or tab
        if (this.arrow_initial) {
          this._arrowSelectionCount = (this._arrowSelectionCount + 1) % peopleList.length;
        } else {
          this.arrow_initial = true;
          this._arrowSelectionCount = 0;
        }
      }

      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.length; i++) {
        peopleList[i].classList.remove('suggestion-focused');
      }

      // set selected background
      const focusedItem = peopleList[this._arrowSelectionCount];
      if (focusedItem) {
        focusedItem.classList.add('suggestion-focused');
        focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'start' });
      }
    }
  }

  private clearArrowSelection(): void {
    this.arrow_initial = false;
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    for (let i = 0; i < peopleList.length; i++) {
      peopleList[i].classList.remove('suggestion-focused');
    }
  }

  private setEntityConfig() {
    // alias Map map customer entity name to real entity name in suggestion API
    var aliasMap = new Map<string, string>();
    var suggestionConfig: SuggestionConfig = {
      configMap: new Map<String, SuggestionEntityConfig>(),
      queryString: this.input.value
    };
    var suggestionEntityOrder = [];
    // allow rename entity type.
    aliasMap.set('query', 'text');
    var defaultCount = 3;
    var defaultSegment = '-';
    // user input entityTypes should like "query, file-1, people,...other future entity types "
    //
    var entityTypes = this._entityTypes.split(',');
    for (var entityType of entityTypes) {
      var entityConfig = entityType.split(defaultSegment);
      //remove spaces and to lower case
      for (var key in entityConfig) {
        entityConfig[key] = entityConfig[key].toLowerCase().trim();
      }

      var maxCount = defaultCount;
      if (entityConfig.length > 1) {
        maxCount = parseInt(entityConfig[1]);
      }
      var suggestionEntityConfig: SuggestionEntityConfig = {
        maxCount: maxCount
      };

      var entityName = aliasMap.has(entityConfig[0]) ? aliasMap.get(entityConfig[0]) : entityConfig[0];
      suggestionConfig.configMap.set(
        aliasMap.has(entityConfig[0]) ? aliasMap.get(entityConfig[0]) : entityConfig[0],
        suggestionEntityConfig
      );

      suggestionEntityOrder.push(entityName);
    }
    // set entity render order
    this._suggestionConfig = suggestionConfig;
    // set entity query config
    this._suggestionEntityOrder = suggestionEntityOrder;
  }

  private getFocusItemValue() {
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    for (let i = 0; i < peopleList.length; i++) {
      if (peopleList[i].classList.contains('suggestion-focused')) {
        return this._suggestionItemsMap.get(peopleList[i].id);
      }
    }
  }

  /**
   * Render the flyout chrome.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearch
   */
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    return html`
        <mgt-flyout light-dismiss class="flyout">
          ${anchor}
          <div slot="flyout" class="flyout-root">
            ${this.renderFlyoutContent()}
          </div>
        </mgt-flyout>
      `;
  }

  /**
   * Render the appropriate state in the results flyout.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearch
   */
  protected renderFlyoutContent(): TemplateResult {
    if (this.isLoadingState || this._showLoading) {
      return this.renderLoading();
    }

    if (this.isEmptySuggestion()) {
      return this.renderNoData();
    }

    return html`
          ${this.renderFlyoutHeader()}
          ${repeat(this._suggestionEntityOrder, entityType => {
            return this.renderEntityRouter(entityType);
          })}
          ${this.renderFlyoutFooter()}
  
      `;

    // render query result
    // render file
  }

  protected renderFlyoutHeader() {
    return this.renderTemplate('flyout-header', {}) || html``;
  }

  protected renderFlyoutFooter() {
    return this.renderTemplate('flyout-footer', {}) || html``;
  }

  protected renderEntityRouter(entityType: string) {
    var data = this._foundSuggestion.get(entityType);
    if (entityType == 'file') {
      return this.renderFileSearchResults(data);
    }
    if (entityType == 'text') {
      return this.renderQuerySearchResults(data);
    }
    if (entityType == 'people') {
      return this.renderPeopleSearchResults(data);
    }
    return this.renderCustomizedSearchResults(entityType, data);
  }

  // render People search result
  protected renderPeopleSearchResults(people?: SuggestionPeople[]) {
    if (people == null || people == undefined) return html``;
    const input = this.userInput;
    return html`
        ${this.renderPeopleHeader()}
        ${this.renderPeople(people)}
    `;
  }

  // render File search result
  protected renderFileSearchResults(files?: SuggestionFile[]) {
    if (files == null || files == undefined) return html``;
    if (files.length < 1) return html``;
    const input = this.userInput;
    return html`
        ${this.renderFileHeader()}
        ${this.renderFiles(files)}
    `;
  }

  // render Query search result
  protected renderQuerySearchResults(querys?: SuggestionQuery[]) {
    if (querys == null || querys == undefined) return html``;
    if (querys.length < 1) return html``;
    const input = this.userInput;
    return html`
        ${this.renderQueryHeader()}
        <div>
        ${this.renderQuerys(querys)}
        </div>

    `;
  }

  // Rendering entity header
  protected renderPeopleHeader() {
    return (
      this.renderTemplate('suggestion-people-header', null) ||
      html`
      <div class="suggestion-entity-label">
        People
      </div>
      `
    );
  }

  protected renderQueryHeader() {
    return (
      this.renderTemplate('suggestion-query-header', {}) ||
      html`
      <div class="suggestion-entity-label">
        Suggested Searches
      </div>
      `
    );
  }

  protected renderFileHeader() {
    return (
      this.renderTemplate('suggestion-file-header', {}) ||
      html`
      <div class="suggestion-entity-label">
        Files
      </div>
      `
    );
  }

  protected renderCustomizedHeader(entityType: string) {
    console.log('customized-' + entityType + '-header');
    return this.renderTemplate('customized-' + entityType + '-header', null) || html``;
  }

  // rendering Querys / people / files / customized entities

  protected renderQuerys(querys?: SuggestionQuery[]) {
    querys = querys || this._foundQuerySuggestion;
    if (querys.length < 1) return html``;

    const input = this.userInput;
    return html`
      ${repeat(querys, query => {
        return html`
          <div
            class="suggestion-query-container suggestion-common-container"
            id="${query.referenceId}"
            @click="${e => this.fireCustomEvent('onEntityClick', query)}"
          >
          ${this.renderQuery(query)}
          </div>

        `;
      })}
    `;
  }

  protected renderPeople(people?: SuggestionPeople[]) {
    return html`
      ${repeat(people, person => {
        return html`
        <div
            class="suggestion-common-container"
            id="${person.referenceId}"
            @click="${e => {
              this.fireCustomEvent('suggestionClick', person);
            }}"
          >
          ${this.renderPerson(person)}

        </div>
        `;
      })}
    `;
  }

  protected renderFiles(files?: SuggestionFile[]) {
    return html`
      ${repeat(files, file => {
        return html`
          <div
            class="suggestion-common-container"
            id="${file.referenceId}"
            @click="${e => {
              this.fireCustomEvent('suggestionClick', file);
            }}"
          >
          ${this.renderFile(file)}

          </div>
        `;
      })}
    `;
  }

  // render customized search results
  protected renderCustomizedSearchResults(entityType: string, data) {
    return html`
      ${this.renderCustomizedHeader(entityType)}
      ${this.renderCustomizedEntities(entityType, data)}
      `;
  }

  protected renderCustomizedEntities(entityType: string, data: any[]) {
    if (data.length < 1) return html``;
    const input = this.userInput;
    return html`
      ${repeat(data, ele => {
        return html`
          <div
            class="suggestion-common-container"
            id="${ele.referenceId}"
            @click="${e => this.fireCustomEvent('suggestionClick', ele)}"
          >
          ${this.renderCustomizedEntity(entityType, ele)}
          </div>

        `;
      })}
    `;
  }

  // render single query / file / person / customized entity

  protected renderQuery(query?: SuggestionQuery) {
    const input = this.userInput;
    return (
      this.renderTemplate('suggested-query', query, query.referenceId) ||
      html`
      <div class="suggestion-content-container">
              <div class="suggestion-query-description">
                <b>${query.query.slice(0, input.length)}</b><span>${query.query.slice(input.length)}</span>
              </div>
      </div>

      `
    );
  }

  protected renderPerson(person?: SuggestionPeople) {
    const input = this.userInput;
    return (
      this.renderTemplate('suggested-person', person, person.referenceId) ||
      html`
          <mgt-person
            .personDetails=${this.ConvertSuggestionPersonToIDynamicPerson(person)}
            .fetchImage=${true}
            class="mgt-suggestion-person-default"
            view="threeLines"
          ></mgt-person>
      `
    );
  }

  protected renderFile(file?: SuggestionFile) {
    const input = this.userInput;
    return (
      this.renderTemplate('suggested-file', file, file.referenceId) ||
      html`
          <mgt-file
            .fileDetails=${this.ConvertSuggestionFileToDriveItem(file)}
            class="mgt-suggestion-file-default"
          ></mgt-file>

      `
    );
  }

  protected renderCustomizedEntity(entityType: string, data: any) {
    return this.renderTemplate('customized-' + entityType, data, data.referenceId) || html``;
  }

  // Covert Suggestion Entity to File/Query/People

  private ConvertSuggestionFileToDriveItem(file: SuggestionFile): DriveItem {
    let driveItem: DriveItem = {
      size: file.FileSize,
      webDavUrl: file.AccessUrl,
      fileSystemInfo: {
        lastModifiedDateTime: file.DateModified
      },
      name: file.name
    };
    return driveItem;
  }

  private ConvertSuggestionPersonToIDynamicPerson(person: SuggestionPeople): IDynamicPerson {
    let dynamicPerson: IDynamicPerson = {
      displayName: person.displayName,
      jobTitle: person.jobTitle,
      personImage: person.personImage,
      imAddress: person.imAddress
    };
    return dynamicPerson;
  }

  private isEmptySuggestion(): Boolean {
    if (this._foundSuggestion == undefined) return true;
    for (var suggestion of this._foundSuggestion) {
      if (suggestion != null && suggestion.length > 0) {
        return false;
      }
    }
    return true;
  }
}
