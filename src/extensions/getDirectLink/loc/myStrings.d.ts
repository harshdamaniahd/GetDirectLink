declare interface IGetDirectLinkCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'GetDirectLinkCommandSetStrings' {
  const strings: IGetDirectLinkCommandSetStrings;
  export = strings;
}
