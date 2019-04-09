#I __SOURCE_DIRECTORY__
#I @"../packages"
#r @"NETStandard.Library.NETFRamework/build/net461/lib/netstandard.dll"
#r @"DocumentFormat.OpenXml/lib/net46/DocumentFormat.OpenXml.dll"
#r @"System.IO.Packaging/lib/net46/System.IO.Packaging.dll"
#r @"WindowsBase"
#r @"../src/sales-tools-library/bin/debug/netstandard2.0/sales-tools-library.dll"


open Design.SalesTools.Sei.Import 



let takeoffFileName = @"C:\Users\darien.shannon\code\DESign\sales-tools\test-work-books\test-sei-takeoff.xlsm"


let printJoists () =
    use bom = GetTakeoff takeoffFileName
    let joists = GetJoists bom
    joists
    |> Seq.iter
        (fun joist -> printfn "%A" joist)


printJoists()
