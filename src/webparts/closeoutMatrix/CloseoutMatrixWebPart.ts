import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import CloseoutMatrix, { ICloseoutMatrixProps } from "./components/CloseoutMatrix";

export interface ICloseoutMatrixWebPartProps {
  listTitle: string;
}

export default class CloseoutMatrixWebPart extends BaseClientSideWebPart<ICloseoutMatrixWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICloseoutMatrixProps> = React.createElement(CloseoutMatrix, {
      context: this.context,
      listTitle: this.properties.listTitle || "CLOCloseoutMatrix"
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Closeout Matrix Settings"
          },
          groups: [
            {
              groupName: "Data",
              groupFields: [
                PropertyPaneTextField("listTitle", {
                  label: "SharePoint List title",
                  value: this.properties.listTitle
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

