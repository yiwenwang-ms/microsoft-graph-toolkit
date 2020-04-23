import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';

import * as strings from 'MgtDemoWebPartStrings';
import MgtDemo from './components/MgtDemo';
import { IMgtDemoProps } from './components/IMgtDemoProps';

// import polyfills for ie11
import 'regenerator-runtime';
import 'core-js/features/array';
import 'core-js/features/array-buffer';
import 'core-js/es/object';
import 'core-js/es/function';
import 'core-js/es/parse-int';
import 'core-js/es/parse-float';
import 'core-js/es/number';
import 'core-js/es/math';
import 'core-js/es/string';
import 'core-js/es/date';
import 'core-js/es/array';
import 'core-js/es/regexp';

// import web component polyfills for browsers that need them
import '@webcomponents/webcomponentsjs/webcomponents-bundle.js';

// import the providers at the top of the page
import { Providers, SharePointProvider } from '../../../../../dist/commonjs';

export interface IMgtDemoWebPartProps {
  description: string;
}

export default class MgtDemoWebPart extends BaseClientSideWebPart<IMgtDemoWebPartProps> {
  // set the global provider
  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IMgtDemoProps> = React.createElement(MgtDemo, {
      description: this.properties.description
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
