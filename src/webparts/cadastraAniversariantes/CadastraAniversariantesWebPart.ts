import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'CadastraAniversariantesWebPartStrings';


import { sp, ItemAddResult } from "@pnp/sp";
import * as $ from "jquery";
import "bootstrap";

require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

//css padrao
require('../../stylelibrary/css/padrao.css');

//importa toast
require('../../stylelibrary/css/toastr.min.css');
import toastr from '../../stylelibrary/js/toast/toastr.min.js';

export interface ICadastraAniversariantesWebPartProps {
  description: string;
}

export default class CadastraAniversariantesWebPart extends BaseClientSideWebPart<ICadastraAniversariantesWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {

    //carrego o template de layout
    this.domElement.innerHTML = require("./template.html");

    document
      .getElementById("btnSalvar")
      .addEventListener("click", () => this.SalvarAniversariante());

    this.ObterTodasAreas();
  }

  public async ObterTodasAreas() {

    let montahtmlAreas;

    await sp.web.lists
      .getByTitle("Areas2")
      .items.top(5000)
      .select("ID, Title")
      .orderBy("Title",true)
      .get()
      .then(items => {

        montahtmlAreas = `<option value="">Selecione</option>`;
        items.forEach(element => {
          montahtmlAreas += `<option value="${element.Title}">${element.Title}</option>`;
        });

        $("#ddlArea").html(montahtmlAreas);

      })
      .catch(e => {
        console.log("erro", e);
      });
  }

  protected SalvarAniversariante(): void {

    toastr.options = {
      "closeButton": false,
      "debug": false,
      "newestOnTop": false,
      "progressBar": true,
      "positionClass": "toast-top-center",
      "preventDuplicates": false,
      "onclick": null,
      "showDuration": "300",
      "hideDuration": "1000",
      "timeOut": "4000",
      "extendedTimeOut": "1000",
      "showEasing": "swing",
      "hideEasing": "linear",
      "showMethod": "fadeIn",
      "hideMethod": "fadeOut"
    };

    let Nome = $("#txtNome").val();
    let DataAniversario = new Date(`${$("#txtDataAniversario").val()}T12:00:00`).toISOString();
    let Area = $("#ddlArea option:selected").text();
    let UrlFoto = $("#txtURLFoto").val();
    let Observacao = $("#txtObservacao").val();

    sp.web.lists.getByTitle("Aniversariantes2").items.add({
      Title: Nome,
      DataAniversario: DataAniversario,
      Area: Area,
      UrlFoto: UrlFoto,
      Observacao: Observacao
    }).then((iar: ItemAddResult) => {
      toastr["success"]("Cadastro realizado com sucesso!", "Sucesso");
    },
      (err) => {
        console.log(err);
        toastr["error"]("Ocorreu um erro ao realizar o cadastro.", "Erro");
      });
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
