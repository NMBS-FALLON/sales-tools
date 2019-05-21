open OfficeOpenXml.Helpers
module OfficeOpenXml.Helpers

open System
open System.Text
open OfficeOpenXml
open OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime

let ExcelColNameToInt (excelName : string) =
    Encoding.ASCII.GetBytes(excelName)
    |> Seq.rev
    |> Seq.mapi
        (fun i b ->
            let letterPos = (float)b - 64.0
            let value = letterPos * (26.0**(float i))
            int value)
    |> Seq.sum

let TryGetValueAsString (worksheet: ExcelWorksheet) row columnName =
    let possibleCellString = worksheet.GetValue<string>(row, ExcelColNameToInt columnName)
    match possibleCellString with
    | null -> None
    | s when s.Trim() = "" -> None
    | _ -> Some possibleCellString

(*
let inline TryGetValue< ^T when ^T : (static member Parse : string -> ^T)> (worksheet: ExcelWorksheet) row columnName =
  let stringValue = TryGetValueAsString worksheet row columnName
   
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
*)

let inline TryGetValue< ^T> ((worksheet : ExcelWorksheet), row, columnName) =
  let valueAsString = TryGetValueAsString worksheet row columnName
   
  let result =
      match valueAsString with
      | Some s ->
        try
          if typeof< ^T> = typeof<string> then
            (box s) :?> ^T |> Some |> Ok
          else
            let o = Activator.CreateInstance< ^T>()
            match box o with
            | :? int -> (box (FSharp.Core.int.Parse s)) :?> ^T |> Some |> Ok
            | :? float -> (box (FSharp.Core.float.Parse s)) :?> ^T |> Some |> Ok
            | :? bool -> (box (FSharp.Core.bool.Parse s)) :?> ^T |> Some |> Ok
            | :? DateTime -> (box (DateTime.Parse s)) :?> ^T |> Some |> Ok
        with
          _ -> 
            sprintf "Failed while attempting to parse %s as a %A at %s"
              (worksheet.GetValue<string>(row, ExcelColNameToInt columnName))
              typeof< ^T>
              (sprintf "%s%i" columnName row)
            |> Error
      | None -> Ok None
  result




