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
        |> Seq.where (fun sheet -> sheet.Name.ToUpper().Contains("JST.") || sheet.Name.ToUpper().Contains("TAKEOFF"))
        |> Seq.tryHead

    match possibleSheet with
    | Some sheet -> Ok sheet
    | None -> Error "Takeoff does not contain a sheet with the name 'Jst.TO' or 'Takeoff', please rename the sheet with the takeoff and re-try."

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
               { JoistDto.Mark = (takeoff.TakeoffSheet, row, "A") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.Quantity = (takeoff.TakeoffSheet, row, "B") |> TryGetValue<int> |> handleWithFailure
                 JoistDto.Depth = (takeoff.TakeoffSheet, row, "C") |> TryGetValue<float> |> handleWithFailure
                 JoistDto.Series = (takeoff.TakeoffSheet, row, "D") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.Designation = (takeoff.TakeoffSheet, row, "E") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.BaseLength = (takeoff.TakeoffSheet, row, "F") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.Slope = (takeoff.TakeoffSheet, row, "R") |> TryGetValue<float> |> handleWithFailure
                 JoistDto.PitchType = (takeoff.TakeoffSheet, row, "S") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.SeatDepthLeft = (takeoff.TakeoffSheet, row, "U") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.SeatDepthRight = (takeoff.TakeoffSheet, row, "V") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.TcxlLength = (takeoff.TakeoffSheet, row, "AC") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.TcxlType = (takeoff.TakeoffSheet, row, "AB") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.TcxrLength = (takeoff.TakeoffSheet, row, "AD") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.TcxrType = (takeoff.TakeoffSheet, row, "AE") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.AddLoad =
                    let asString = (takeoff.TakeoffSheet, row, "AH") |> TryGetValue<string> |> handleWithFailure
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.NetUplift = (takeoff.TakeoffSheet, row, "AI") |> TryGetValue<float> |> handleWithFailure
                 JoistDto.AxialLoad =
                    let asString = (takeoff.TakeoffSheet, row, "AK") |> TryGetValue<string> |> handleWithFailure
                    asString |> Option.map (fun v -> FSharp.Core.float.Parse (v.Replace("K", "").Replace("k", "")))
                 JoistDto.AxialType = (takeoff.TakeoffSheet, row, "AL") |> TryGetValue<string> |> handleWithFailure
                 JoistDto.Notes = (takeoff.TakeoffSheet, row, "AT") |> TryGetValue<string> |> handleWithFailure } ]
        
        


    joistDtos
    |> Seq.filter (fun jdto -> jdto.Quantity.IsSome)
    |> Seq.map Joist.Parse
    |> Seq.mapi (fun i j -> {j with Id = i + 1})




