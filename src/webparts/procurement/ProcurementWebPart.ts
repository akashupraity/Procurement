import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'ProcurementWebPartStrings';
import Procurement from './components/Procurement';
import { IProcurementProps } from './components/IProcurementProps';

export interface IProcurementWebPartProps {
  description: string;
}

export default class ProcurementWebPart extends BaseClientSideWebPart<IProcurementWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });  
    });
  }
  public render(): void {
    const element: React.ReactElement<IProcurementProps> = React.createElement(
      Procurement,
      {
        //description: this.properties.description
        siteUrl:this.context.pageContext.web.absoluteUrl,
        context:this.context,
        procurementRequestList:"ProcurementRequests",
        procurementReqDetailList:"ProcurementRequestDetails",
        logHistoryListTitle:"ProcurementLogHistory",
      }
    );

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
