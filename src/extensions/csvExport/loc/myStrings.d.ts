declare interface ICsvExportCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CsvExportCommandSetStrings' {
  const strings: ICsvExportCommandSetStrings;
  export = strings;
}
