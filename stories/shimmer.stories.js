/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';
import '../dist/es6/components/common/mgt-shimmer/mgt-shimmer';

export default {
  title: 'mgt-shimmer',
  component: 'mgt-shimmer',
  decorators: [withA11y, withCodeEditor]
};

export const shimmer = () => html`
  <mgt-shimmer></mgt-shimmer>
`;

export const shimmer2 = () => html`
  <mgt-shimmer>
    <mgt-shimmer-circle></mgt-shimmer-circle>
    <mgt-shimmer-gap width="2%"></mgt-shimmer-gap>
    <mgt-shimmer-circle height="15" vertical-align="top"></mgt-shimmer-circle>
    <mgt-shimmer-gap width="2%"></mgt-shimmer-gap>
    <mgt-shimmer-line vertical-align="bottom" width="20%"></mgt-shimmer-line>
    <mgt-shimmer-gap></mgt-shimmer-gap>
    <mgt-shimmer-line vertical-align="top" width="20%" height="5"></mgt-shimmer-line>
    <mgt-shimmer-gap></mgt-shimmer-gap>
    <mgt-shimmer-line width="15%" height="16"></mgt-shimmer-line>
    <mgt-shimmer-gap></mgt-shimmer-gap>
    <mgt-shimmer-line vertical-align="bottom" height="10"></mgt-shimmer-line>
  </mgt-shimmer>
`;
