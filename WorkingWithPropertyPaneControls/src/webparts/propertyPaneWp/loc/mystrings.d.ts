declare interface IPropertyPaneWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  discount: number;
  billAmount: number;
  netBillAmount: number;
}

declare module 'PropertyPaneWpWebPartStrings' {
  const strings: IPropertyPaneWpWebPartStrings;
  export = strings;
}
