module Design.SalesTools.Gem.CreateTakeoff
open System.IO
open System.IO

open System
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open System.IO
open System.Reflection

[<Measure>] type inch
[<Measure>] type ft

let inchToFoot (inch : float<inch>) = inch / 12.0<inch/ft>
let footToInch (ft : float<ft>) = ft * 12.0<inch/ft>

type Joist =
    {
    Quantity : int
    Description : string
    BaseLengthFt : float<ft>
    BaseLengthIn : float<inch>
    BCExtension : bool
    PitchType : string
    Slope : float<inch/ft>
    OriginalMark : string
    }

let nullableToOption<'T> value =
    match (box value) with
    | null  -> None
    | value when value = (box "") -> None
    | _ -> Some ((box value) :?> 'T)

let CreateTakeoff gemTakeoffFileName = 
    #if INTERACTIVE
    #r "../packages/Deedle.1.2.5/lib/net40/Deedle.dll"
    #r "Microsoft.Office.Interop.Excel.dll"
    #r "../packages/FSharp.Configuration.1.3.0/lib/net45/FSharp.Configuration.dll"
    System.Environment.CurrentDirectory <- @"C:\Users\darien.shannon\Documents\Code\F#\FSharp\NMBS_TOOLS\NMBS_TOOLS\bin\Debug"
    #endif
    
    let getInfoFunction (workBook: Workbook) =
        let takeoff = workBook.Worksheets.["Takeoff"] :?> Worksheet
        let H3Text = takeoff.Range("H3").Value2 :?> string

        let takeoffArray =
            if H3Text.ToUpper().Contains("BASE") then
                takeoff.Range("C4", "M3000").Value2 :?> obj[,]
            else
                takeoff.Range("A4", "K3000").Value2 :?> obj[,]

        let startRow = Array2D.base1 takeoffArray
        let endRow = 
            match startRow with
            | 0 -> (Array2D.length1 takeoffArray) - 1
            | _ -> (Array2D.length1 takeoffArray)

        let startColumn = Array2D.base2 takeoffArray

        let splitBaseLength baseLength =
             let baseLength = if baseLength = null then "" else baseLength
             let baseLengthArray = baseLength.Split([|"."|], StringSplitOptions.RemoveEmptyEntries)
             let baseLengthFt = if baseLengthArray.Length < 1 then 0.0 else float (baseLengthArray.[0])
             let baseLengthIn = if baseLengthArray.Length < 2 then 0.0
                                else
                                    if baseLengthArray.[1] = "1" then 10.0
                                    else float (baseLengthArray.[1])
             (baseLengthFt * 1.0<ft>, baseLengthIn * 1.0<inch>)


        [for row = startRow to endRow do
             let (_, quantity) = Int32.TryParse(Convert.ToString(takeoffArray.[row, startColumn + 1]))
               
             if quantity <> 0 then
                 let originalMark = string (takeoffArray.[row, startColumn])
                 let depth = Convert.ToString(takeoffArray.[row, startColumn + 2])
                 let series = Convert.ToString(takeoffArray.[row, startColumn + 3 ])
                 let designation = Convert.ToString(takeoffArray.[row, startColumn + 4])
                 let description = (depth + series + designation).Replace(" ", "")
                 let baseLength = Convert.ToString(takeoffArray.[row, startColumn + 5])
                 let baseLengthFt = fst (splitBaseLength baseLength)
                 let baseLengthIn = snd (splitBaseLength baseLength)
                 let bcExtension = Convert.ToString(takeoffArray.[row, startColumn + 8]).Contains("B")
                 let pitchType = Convert.ToString(takeoffArray.[row, startColumn + 9]).Trim()
                 let slope = 
                     let slopeValue = takeoffArray.[row, startColumn + 10]
                     if slopeValue = null || slopeValue = (box "") then 0.0<inch/ft>
                     else (Double.Parse (string slopeValue)) * 1.0<inch/ft>
                 yield
                     {
                     Quantity = quantity
                     Description = description
                     BaseLengthFt = baseLengthFt
                     BaseLengthIn = baseLengthIn
                     BCExtension = bcExtension
                     PitchType = pitchType
                     Slope = slope
                     OriginalMark = originalMark
                     }]

    let getAllInfo (reportPath : string) (getInfoFunction : Workbook -> 'TOutput list) =
        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        let info =
            let tempReportPath = System.IO.Path.GetTempFileName()
            File.Delete(tempReportPath)
            File.Copy(reportPath, tempReportPath)
            let workbook = tempExcelApp.Workbooks.Open(tempReportPath)
            let info = getInfoFunction workbook
            workbook.Close(false)
            Marshal.ReleaseComObject(workbook) |> ignore
            System.GC.Collect() |> ignore
            printfn "Finished processing %s." reportPath
            info
        tempExcelApp.Quit()
        Marshal.ReleaseComObject(tempExcelApp) |> ignore
        System.GC.Collect() |> ignore
        info

    let inputAllInfo (joists : Joist list) (savePath: string) =
        printfn "Creating NMBS Takeoff; Please hold."
        let tempExcelApp = Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        tempExcelApp.AutomationSecurity <- Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        tempExcelApp.DisplayAlerts <- false
        try
            let excelPath = System.IO.Path.GetTempFileName();

            let assembly = Assembly.GetExecutingAssembly()
            let blankSalesBomResourceName = "sales_tools.resources.BLANK SALES BOM.xlsm"
            use blankSalesBomStream = assembly.GetManifestResourceStream(blankSalesBomResourceName)
            use blankSalesBomMemoryStream = new MemoryStream()
            blankSalesBomStream.CopyTo(blankSalesBomMemoryStream)
            System.IO.File.WriteAllBytes(excelPath, blankSalesBomMemoryStream.ToArray())

            let bom = tempExcelApp.Workbooks.Open(excelPath)

            let addJoistSheet index name =                                       
                let blankJoistSheet = bom.Worksheets.["J(BLANK)"] :?> Worksheet     
                blankJoistSheet.Copy(bom.Worksheets.[index])
                let newJoistSheet = (bom.Worksheets.[index]) :?> Worksheet
                newJoistSheet.Name <- name
                newJoistSheet

            let coverIndex = (bom.Worksheets.["Cover"] :?> Worksheet).Index

            let mutable pageCount = 1
            let mutable joistSheet = addJoistSheet (coverIndex + pageCount) (sprintf "J (%i)" pageCount)

            let mutable row = 6
            let mutable markCount = 1
            for joist in joists do
                if row > 41 then
                    pageCount <- pageCount + 1
                    joistSheet <- addJoistSheet (coverIndex + pageCount) (sprintf "J (%i)" pageCount)
                    row <- 6
                let sRow = row.ToString()
                joistSheet.Range("A" + sRow).Value2 <- markCount.ToString()
                joistSheet.Range("B" + sRow).Value2 <- joist.Quantity
                joistSheet.Range("C" + sRow).Value2 <- joist.Description
                joistSheet.Range("Z" + sRow).Value2 <- (sprintf "Original Mark: %s" joist.OriginalMark)

                if (joist.PitchType = null || joist.PitchType = "") then
                    let run = joist.BaseLengthFt + inchToFoot joist.BaseLengthIn
                    let rise = inchToFoot (run * joist.Slope)
                    let slopeLength = sqrt (run*run + rise*rise)
                    let slopeLengthFt = slopeLength - (slopeLength % 1.0<ft>)
                    let slopeLengthIn =  footToInch (slopeLength % 1.0<ft>)
                    joistSheet.Range("D" + sRow).Value2 <- slopeLengthFt
                    if slopeLengthIn % 1.0<inch> > 0.2<inch> then
                        joistSheet.Range("E" + sRow).Value2 <- Math.Ceiling(float slopeLengthIn)
                    else
                        joistSheet.Range("E" + sRow).Value2 <- Math.Floor(float slopeLengthIn)
                else
                    joistSheet.Range("D" + sRow).Value2 <- joist.BaseLengthFt
                    joistSheet.Range("E" + sRow).Value2 <- joist.BaseLengthIn

                if joist.BCExtension then
                    joistSheet.Range("N" + sRow).Value2 <- joist.Quantity * 2                  

                row <- row + 3
                markCount <- markCount + 1

            let blankJoistSheet = bom.Worksheets.["J(BLANK)"] :?> Worksheet
            blankJoistSheet.Delete()
            

            let savePath =
                let savePath = savePath.ToUpper()
                savePath.Substring(0, savePath.IndexOf(".XL")) + " (IMPORT).xlsm"

            bom.SaveAs(savePath)

            bom.Close()
            Marshal.ReleaseComObject(bom) |> ignore
            System.GC.Collect() |> ignore


        finally
            tempExcelApp.Quit()
            Marshal.ReleaseComObject(tempExcelApp) |> ignore
            System.GC.Collect() |> ignore    

    inputAllInfo (getAllInfo gemTakeoffFileName getInfoFunction) gemTakeoffFileName

