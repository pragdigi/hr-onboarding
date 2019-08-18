import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./OnboardingHrWebPart.module.scss";
import * as strings from "OnboardingHrWebPartStrings";
import * as microsoftTeams from "@microsoft/teams-js";


export interface IOnboardingHrWebPartProps {
  description: string;
}

export default class OnboardingHrWebPart extends BaseClientSideWebPart<
  IOnboardingHrWebPartProps
> {
  private teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      if (this.context.microsoftTeams) {
        this.context.microsoftTeams.getContext(context => {
          this.teamsContext = context;
          resolve();
        });
      } else {
        resolve();
      }
    });
  }

  public render(): void {
    let title: string = this.teamsContext ? "Teams" : "SharePoint";
    let currentLocation: string = this.teamsContext
      ? `Team: ${this.teamsContext.teamName}`
      : `site collection ${this.context.pageContext.web.title}`;
    this.domElement.innerHTML = `
      <div class="${styles.onboardingHr}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to HR Team - ${title}</span>
              <p class="${
                styles.subTitle
              }">This could have infomration about onboarding. We are located in: ${currentLocation}</p>
              <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
