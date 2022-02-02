import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import * as strings from "ListaAniversariantesWebPartStrings";

import * as $ from "jquery";
import "bootstrap";

require("../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
//css padrao
require("../../stylelibrary/css/padrao.css");

export interface IListaAniversariantesWebPartProps {
  description: string;
}

interface OptionProps {
  month?: any;
  day?: any;
  year?: any;
}

export default class ListaAniversariantesWebPart extends BaseClientSideWebPart<IListaAniversariantesWebPartProps> {
  public ListaAniversariantes() {
    const option: OptionProps = {
      month: "long",
      day: "numeric",
      year: "numeric",
    };

    $.ajax({
      url:
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getByTitle('Aniversariantes2')/items?$select=ID,Title,DataAniversario,Area,UrlFoto`,
      method: "GET",
      async: false,
      headers: {
        Accept: "application/json; odata=verbose",
      },
      success: (data) => {
        let html = `<div class="row"><div class="col-md-12">Nenhum aniversariante hoje</div></div>`;

        if (data.d.results.length > 0) {
          html = "";
          $.each(data.d.results, (i, result) => {
            html +=
              `<div class="row mt-3">` +
              `<div class="col-md-2"><img class="foto" src="${result.UrlFoto}"/></div>` +
              `<div class="col-md-3">${result.Title}</div>` +
              `<div class="col-md-3">${new Date(
                result.DataAniversario
              ).toLocaleDateString("pt-br", option)}</div>` +
              `</div>`;
          });

          $("#divAniversariantes").html(html);
        } else {
          $("#divAniversariantes").html(html);
        }
      },
      error: (errorCode, errorMessage) => {
        console.log(
          "Erro ao recuperar os itens. \nError: " +
            errorCode +
            "\nStackTrace: " +
            errorMessage
        );
      },
    });
  }

  public render(): void {
    //carrego o template de layout
    this.domElement.innerHTML = require("./template.html");
    $("#lblTitulo").html(`${escape(this.properties.description)}`);
    this.ListaAniversariantes();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
