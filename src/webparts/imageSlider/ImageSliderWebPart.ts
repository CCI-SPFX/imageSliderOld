import * as React from 'react';
import * as ReactDom from 'react-dom';
import "@pnp/sp/webs";
import "@pnp/sp/lists";

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ImageSlider from './components/ImageSlider';
import { IImageSliderProps } from './components/IImageSliderProps';

import { spfi, SPFx, SPFI } from "@pnp/sp";

export interface IImageSliderWebPartProps {
  selectedLibrary: string;
}

export default class ImageSliderWebPart extends BaseClientSideWebPart<IImageSliderWebPartProps> {

  private _sp!: SPFI;
  private _libraries: { key: string; text: string }[] = [];
 
  protected async onInit(): Promise<void> {
    await super.onInit();

    this._sp = spfi().using(SPFx(this.context));

    await this.loadLibraries();
  }
 
  private async loadLibraries(): Promise<void> {
    try {
      const lists = await this._sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false")
        .select("Title")();

      this._libraries = lists.map(l => ({
        key: l.Title,
        text: l.Title
      }));

      // refresh property pane for dropdown updates
      this.context.propertyPane.refresh();
    } catch (err) {
      console.error(err);
    }
  }
 
  public render(): void {

    const element: React.ReactElement<IImageSliderProps> = React.createElement(
      ImageSlider,
      {
        sp: this._sp,
        selectedLibrary: this.properties.selectedLibrary
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
            description: "Image Slider Settings"
          },
          groups: [
            { 
              groupFields: [
                PropertyPaneDropdown("selectedLibrary", {
                  label: "Select Document Library",
                  options: this._libraries
                })
              ]
            }
          ]
        }
      ]
    };
  }
}