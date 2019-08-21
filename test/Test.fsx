#I __SOURCE_DIRECTORY__
#I @"../packages"
#r @"NETStandard.Library.NETFRamework/build/net461/lib/netstandard.dll"
#r @"EPPlus/lib/net40/EPPlus.dll"
#r @"WindowsBase"
#r @"../src/bin/release/netstandard2.0/sales-tools-library.dll"


open Design.SalesTools.Sei.Import 



let seiTakeoffFileName = @"C:\Users\darien.shannon\code\sales-tools\test-work-books\sei\31733 2019.05.21 RFP Joist.xlsm"


let printJoists () =
    use bom = GetTakeoff seiTakeoffFileName
    let joists = GetJoists bom
    joists
    |> Seq.iter
        (fun joist -> printfn "%A" joist)


Design.SalesTools.Sei.CreateTakeoff.CreateTakeoff(seiTakeoffFileName)