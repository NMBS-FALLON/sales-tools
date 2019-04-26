module Design.SalesTools.Sei.Import 

open System
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open FSharp.Core
open DESign.SpreadSheetML.Helpers
open Design.SalesTools.Sei.Dto   
open Design.SalesTools.Sei.Domain

let handleWithFailwith result =
    match result with
    | Ok v -> v
    | Error s -> failwith s

let getTakeoffSheet (document : SpreadsheetDocument) =
    let possibleSheet =
        document
        |> GetSheetsByPartialName "Jst."
        |> Seq.tryHead
    match possibleSheet with
    | Some sheet -> Ok sheet
    | None -> Error "Takeoff does not contain a sheet with the name 'Jst.TO'"

type Takeoff =
    { Document : SpreadsheetDocument
      StringTable : SharedStringTable
      TakeoffSheet : Worksheet }

    interface IDisposable with
        member this.Dispose() = (this.Document :> IDisposable).Dispose()

    member this.TryGetCellValueAtColumnAsString column row =
        TryGetCellValueAtColumnAsString column this.StringTable row
    member inline this.TryGetCellValueAtColumnWithType column row =
        TryGetCellValueAtColumnWithType column this.StringTable row |> handleWithFailwith

let GetTakeoff(takeoffFileName : string) =
    let document = SpreadsheetDocument.Open(takeoffFileName, true)
    { Document = document
      StringTable = GetStringTable document
      TakeoffSheet = getTakeoffSheet document |> handleWithFailwith}

// let takeoff = GetTakeoff takeoffFileName
let GetJoists(takeoff : Takeoff) =
    let takeoffSheet = takeoff.TakeoffSheet
    let firstRowNum = 5u
    let lastRowNum = 10000u
    let rows = takeoffSheet |> GetRows firstRowNum lastRowNum

    let joistDtos=
        rows
        |> Seq.map
               (fun row ->
               { JoistDto.Mark = row |> takeoff.TryGetCellValueAtColumnAsString "A"
                 JoistDto.Quantity = row |> takeoff.TryGetCellValueAtColumnWithType<int> "B"
                 JoistDto.Depth = row |> takeoff.TryGetCellValueAtColumnWithType<float> "C"
                 JoistDto.Series = row |> takeoff.TryGetCellValueAtColumnAsString "D"
                 JoistDto.Designation = row |> takeoff.TryGetCellValueAtColumnAsString "E"
                 JoistDto.BaseLength = row |> takeoff.TryGetCellValueAtColumnAsString "F"
                 JoistDto.Slope = row |> takeoff.TryGetCellValueAtColumnWithType<float> "R"
                 JoistDto.PitchType = row |> takeoff.TryGetCellValueAtColumnAsString "S"
                 JoistDto.SeatDepthLeft = row |> takeoff.TryGetCellValueAtColumnAsString "U"
                 JoistDto.SeatDepthRight = row |> takeoff.TryGetCellValueAtColumnAsString "V"
                 JoistDto.TcxlLength = row |> takeoff.TryGetCellValueAtColumnAsString "AC"
                 JoistDto.TcxlType = row |> takeoff.TryGetCellValueAtColumnAsString "AB"
                 JoistDto.TcxrLength = row |> takeoff.TryGetCellValueAtColumnAsString "AD"
                 JoistDto.TcxrType = row |> takeoff.TryGetCellValueAtColumnAsString "AE"
                 JoistDto.AddLoad =
                    let asString = row |> takeoff.TryGetCellValueAtColumnAsString "AH"
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.NetUplift = row |> takeoff.TryGetCellValueAtColumnWithType<float> "AI"
                 JoistDto.AxialLoad =
                    let asString = row |> takeoff.TryGetCellValueAtColumnAsString "AK"
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.AxialType = row |> takeoff.TryGetCellValueAtColumnAsString "AL"
                 JoistDto.Notes = row |> takeoff.TryGetCellValueAtColumnAsString "AT" })
        |> Seq.filter (fun joistDto -> not joistDto.IsEmpty)
        
        


    joistDtos
    |> Seq.map Joist.Parse
    |> Seq.mapi (fun i j -> {j with Id = i + 1})




