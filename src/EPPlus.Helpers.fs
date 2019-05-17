module OfficeOpenXml.Helpers

open System
open System.Text
open OfficeOpenXml

let ExcelNameToInt (excelName : string) =
    Encoding.ASCII.GetBytes(excelName)
    |> Seq.rev
    |> Seq.mapi
        (fun i b ->
            let letterPos = (float)b - 64.0
            let value = letterPos * (26.0**(float i))
            int value)
    |> Seq.sum

let TryGetCellValueAtColumnAsString (worksheet: ExcelWorksheet) row columnName =
    let possibleCellString = worksheet.GetValue<string>(row, ExcelNameToInt columnName)
    match possibleCellString with
    | null -> None
    | s when s.Trim() = "" -> None
    | _ -> Some possibleCellString

let inline TryGetCellValueAtColumnWithType< ^T when ^T : (static member Parse : string -> ^T)> (worksheet: ExcelWorksheet) row columnName =
  let stringValue = TryGetCellValueAtColumnAsString worksheet row columnName
   
  let result =
      match stringValue with
      | Some s ->
          try
            match s with
            | _ -> Ok (Some ((^T : (static member Parse : string -> ^T) s)))
          with
            | _ -> Error (sprintf "Failed while attempting to parse %s as a %A at %s" s typeof< ^T> (sprintf "%s%i" columnName row))
      | None -> Ok None
  result