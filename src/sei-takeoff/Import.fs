module Design.SalesTools.Sei.Import 

open System
open FSharp.Core
open Design.SalesTools.Sei.Dto   
open Design.SalesTools.Sei.Domain
open OfficeOpenXml
open OfficeOpenXml.Helpers

let getTakeoffSheet (package : ExcelPackage) =
    let possibleSheet =
        package.Workbook.Worksheets
        |> Seq.where (fun sheet -> sheet.Name.Contains("Jst."))
        |> Seq.tryHead

    match possibleSheet with
    | Some sheet -> Ok sheet
    | None -> Error "Takeoff does not contain a sheet with the name 'Jst.TO'"

type Takeoff =
    { Package : ExcelPackage
      TakeoffSheet : ExcelWorksheet }

    interface IDisposable with
        member this.Dispose() = (this.Package :> IDisposable).Dispose()


     

let GetTakeoff(takeoffFileName : string) =
    let takeoffStream = new System.IO.FileStream(takeoffFileName, System.IO.FileMode.Open)
    let package = new ExcelPackage(takeoffStream)
    { Package = package
      TakeoffSheet = getTakeoffSheet package |> handleWithFailure}


    

// let takeoff = GetTakeoff takeoffFileName
let GetJoists(takeoff : Takeoff) =
    //let takeoff = GetTakeoff(@"C:\Users\darien.shannon\code\sales-tools\test-work-books\broken.xlsm")
    let takeoffSheet = takeoff.TakeoffSheet
    let firstRowNum = 5
    let lastRowNum = 10000

    let joistDtos=
      seq
        [ for row in [firstRowNum..lastRowNum] do
            //let row = 7
            yield
               //let row = rows |> Seq.item 3
               { JoistDto.Mark = (takeoff.TakeoffSheet, row, "A") |||> TryGetCellValueAtColumnAsString
                 JoistDto.Quantity = (takeoff.TakeoffSheet, row, "B") |||> TryGetCellValueAtColumnWithType<int> |> handleWithFailure
                 JoistDto.Depth = (takeoff.TakeoffSheet, row, "C") |||> TryGetCellValueAtColumnWithType<float> |> handleWithFailure
                 JoistDto.Series = (takeoff.TakeoffSheet, row, "D") |||> TryGetCellValueAtColumnAsString
                 JoistDto.Designation = (takeoff.TakeoffSheet, row, "E") |||> TryGetCellValueAtColumnAsString
                 JoistDto.BaseLength = (takeoff.TakeoffSheet, row, "F") |||> TryGetCellValueAtColumnAsString
                 JoistDto.Slope = (takeoff.TakeoffSheet, row, "R") |||> TryGetCellValueAtColumnWithType<float> |> handleWithFailure
                 JoistDto.PitchType = (takeoff.TakeoffSheet, row, "S") |||> TryGetCellValueAtColumnAsString
                 JoistDto.SeatDepthLeft = (takeoff.TakeoffSheet, row, "U") |||> TryGetCellValueAtColumnAsString
                 JoistDto.SeatDepthRight = (takeoff.TakeoffSheet, row, "V") |||> TryGetCellValueAtColumnAsString
                 JoistDto.TcxlLength = (takeoff.TakeoffSheet, row, "AC") |||> TryGetCellValueAtColumnAsString
                 JoistDto.TcxlType = (takeoff.TakeoffSheet, row, "AB") |||> TryGetCellValueAtColumnAsString
                 JoistDto.TcxrLength = (takeoff.TakeoffSheet, row, "AD") |||> TryGetCellValueAtColumnAsString
                 JoistDto.TcxrType = (takeoff.TakeoffSheet, row, "AE") |||> TryGetCellValueAtColumnAsString
                 JoistDto.AddLoad =
                    let asString = (takeoff.TakeoffSheet, row, "AH") |||> TryGetCellValueAtColumnAsString
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.NetUplift = (takeoff.TakeoffSheet, row, "AI") |||> TryGetCellValueAtColumnWithType<float> |> handleWithFailure
                 JoistDto.AxialLoad =
                    let asString = (takeoff.TakeoffSheet, row, "AK") |||> TryGetCellValueAtColumnAsString
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.AxialType = (takeoff.TakeoffSheet, row, "AL") |||> TryGetCellValueAtColumnAsString
                 JoistDto.Notes = (takeoff.TakeoffSheet, row, "AT") |||> TryGetCellValueAtColumnAsString } ]
        
        


    joistDtos
    |> Seq.filter (fun jdto -> jdto.Quantity.IsSome)
    |> Seq.map Joist.Parse
    |> Seq.mapi (fun i j -> {j with Id = i + 1})




