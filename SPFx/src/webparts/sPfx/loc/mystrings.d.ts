declare interface ISPfxStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'sPfxStrings' {
  const strings: ISPfxStrings;
  export = strings;
}
