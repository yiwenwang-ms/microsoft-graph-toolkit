/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-search/mgt-search';

export default {
  title: 'Components | mgt-search',
  component: 'mgt-search',
  decorators: [withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const searchSuggestion = () => html`
  <mgt-search></mgt-search>
`;

export const suggestionCount = () => html`
  <mgt-search entity-types="query-2, file-3, people-1"></mgt-search>
`;

export const suggestionEntity = () => html`
  <mgt-search selected-entity-types="file, text, people"></mgt-search>
`;

export const cvid = () => html`
  <mgt-search cvid="d8e48cff-9cac-40c9-0b5c-e6d24488f781"></mgt-search>
`;

export const textDecorations = () => html`
  <mgt-search text-decorations="1"></mgt-search>
`;

export const customizedEntity = () => html`
<mgt-search entity-types="query,file,people,sample1">
    <template data-type="customized-sample1-header">
        <div>header hahaha</div>
    </template>

    <template data-type="customized-sample1">
        <div>{{this}}</div>
    </template>
</mgt-search>
`;

export const template = () => html`
<mgt-search>
    <template data-type="suggested-query">
        <div>{{this}}</div>
    </template>
</mgt-search>
`;

export const darkTheme = () => html`
  <mgt-search class="mgt-dark"></mgt-search>
`;

export const callback = () =>
  html`
    <mgt-search> </mgt-search>
    <script>


document.querySelector('mgt-search').addEventListener('suggestionClick', e => {
    console.log(e.detail);
});

document.querySelector('mgt-search').addEventListener('enterPress', e => {
  console.log(e.detail);
});


</script>
  `;
