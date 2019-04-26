module Design.SalesTools.Sei.CreateTakeoff
open Design.SalesTools.Sei.Dto
open System
open OfficeOpenXml
open System.IO
open System.Reflection
open OfficeOpenXml
open OfficeOpenXml
open Design.SalesTools.Sei.Import
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Spreadsheet

let CreateTakeoff seiTakeoffFileName =

    use bom = GetTakeoff seiTakeoffFileName
    let joists = GetJoists bom   

    let assembly = Assembly.GetExecutingAssembly()
    let blankSalesBomResourceName = "sales_tools.sei_takeoff.BLANK SALES BOM.xlsm"
    use blankSalesBomStream = assembly.GetManifestResourceStream(blankSalesBomResourceName)
    use package = new ExcelPackage(blankSalesBomStream)
    let mutable rowCount = 1
    let mutable workBookCount = 1
    let mutable ws = package.Workbook.Worksheets.Copy("J(BLANK)", sprintf "J (%i)" workBookCount)
    for joist in joists do
        if rowCount < 38 then
            ()
        else
            workBookCount <- workBookCount + 1
            ws <- package.Workbook.Worksheets.Copy("J(BLANK)", sprintf "J (%i)" workBookCount)
            rowCount <- 1

        let row = rowCount + 4
        
        let baseLengthFt = Math.Floor joist.BaseLength
        let baseLengthIn = Math.Floor ((joist.BaseLength - baseLengthFt) * 12.0)
        ws.SetValue(row, 1, joist.Id)
        ws.SetValue(row, 2, joist.Quantity)
        ws.SetValue(row, 3, joist.Designation)
        ws.SetValue(row, 4, baseLengthFt)
        ws.SetValue(row, 5, baseLengthIn)
        joist.TcxlLength |> Option.iter (fun tcxl -> ws.SetValue(row, 7, tcxl))
        joist.TcxrLength |> Option.iter (fun tcxr -> ws.SetValue(row, 10, tcxr))
        joist.SeatDepthLeft |> Option.iter (fun sdl -> ws.SetValue(row, 12, sdl))
        joist.SeatDepthRight |> Option.iter (fun sdr -> ws.SetValue(row, 13, sdr))
        ws.SetValue(row, 26, (sprintf "Original Mark: %s" joist.TakeoffMark))
        rowCount <- rowCount + 1

    let newTakeoffName = seiTakeoffFileName.ToUpper().Replace(".XLSM", " (IMPORT).XLSM")
    package.SaveAs(FileInfo(newTakeoffName))
   
    


    


