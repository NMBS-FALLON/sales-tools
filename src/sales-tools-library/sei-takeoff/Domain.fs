module Design.SalesTools.Sei.Domain
open Design.SalesTools.Sei.Dto
open System

let handleWithFailure result =
    match result with
    | Ok r -> r
    | Error msg -> failwith msg

type FeetRepresentation =
    FeetDotInch of string
        override this.ToString() =
            match this with
            | FeetDotInch s -> s
        member this.ToFeet() =
            let asString = this.ToString()  
            match asString.Split([|'.'|], StringSplitOptions.RemoveEmptyEntries) with
            | [|ftString|] -> FSharp.Core.float.Parse ftString
            | [|ftString; inString|] ->
                let ftFloat = FSharp.Core.float.Parse ftString
                let inFloat = FSharp.Core.float.Parse inString
                (ftFloat + inFloat / 12.0)
            | _ -> failwith (sprintf "Expecting a 'Feet Dot Inch' but received %s" asString)

type InchRepresentation =
    InchDotFraction of string

        override this.ToString() =
            match this with
            | InchDotFraction s -> s

        member this.ToInch() =
            let asString = this.ToString()
            match asString.Replace("\"","").Split([|'.'|], StringSplitOptions.RemoveEmptyEntries) with
            | [|inString|] -> FSharp.Core.float.Parse inString
            | [|inString; fractionString|] -> 
                let inFloat = FSharp.Core.float.Parse inString
                let fractionFloat =
                    match fractionString.Split([|'/'|], StringSplitOptions.RemoveEmptyEntries) with
                    | [|numeratorString; denominatorString|] ->
                        let numerator = numeratorString |> FSharp.Core.float.Parse
                        let denominator = denominatorString |> FSharp.Core.float.Parse
                        numerator/denominator
                    | _ -> failwith (sprintf "Expecting a 'Inch Dot Fraction' but received %s" asString)
                inFloat + fractionFloat
            | _ -> failwith (sprintf "Expecting a 'Inch Dot Fraction' but received %s" asString)


        
type PitchType =
    | ``AC PP``
    | ``AX PR``
    | BS
    | DP
    |``DP SE``
    | GBL
    | ODP
    | ``ODP SE``
    | PC
    | ``PC SE``
    | ``SC PP``
    | ``SC PR``
    | SP
    | ``SP SE``

    static member Parse pitchType =
        match pitchType with
        | "AC PP" -> ``AC PP`` 
        | "AX PR" -> ``AX PR`` 
        | "BS" -> BS 
        | "DP" -> DP 
        | "DP SE" -> ``DP SE`` 
        | "GBL" -> GBL 
        | "ODP" -> ODP 
        | "ODP SE" -> ``ODP SE`` 
        | "PC" -> PC 
        | "PC SE" -> ``PC SE``
        | "SC PP" -> ``SC PP`` 
        | "SC PR" -> ``SC PR`` 
        | "SP" -> SP 
        | "SP SE" -> ``SP SE`` 
        

type TcxType =
    | S
    | R
    
    static member Parse tcxType =
        match tcxType with
        | "S" -> S
        | "R" -> R

type AxialType =
    | W
    | S
    | WS
    | B

    static member Parse axialType =
        match axialType with
        | "W" -> W
        | "S" -> S
        | "W/S" -> WS
        | "B" -> B

type Load = Load of string

type Joist =
    { Id : int
      TakeoffMark : string
      Quantity : int
      Designation : string
      BaseLength : float
      Slope : float
      PitchType : PitchType option
      SeatDepthLeft : float option
      SeatDepthRight : float option
      TcxlLength : float option
      TcxlType : TcxType option
      TcxrLength : float option
      TcxrType : TcxType option
      Loads : Load seq
      }
    static member Parse (joistDto : JoistDto) =
        let takeoffMark =
            match joistDto.Mark with
            | Some m -> m
            | None -> ""
        let quantity =
            match joistDto.Quantity with
            | Some q -> Ok q
            | None -> Error (sprintf "Mark %s does not have a quantity" takeoffMark)
        let designation =
            match joistDto.Designation with
            | Some d -> d.Replace(" ", "") |> Ok
            | None -> Error (sprintf "Mark %s does not have a section number / load designation" takeoffMark)
        let baseLength = 
            match joistDto.BaseLength with
            Some bl ->
                try
                    FeetDotInch(bl).ToFeet() |> Ok
                with
                | _ -> Error (sprintf "Mark %s's base length is in an improper format" takeoffMark)
            | None -> Error (sprintf "Mark %s does not have a base length" takeoffMark)
        let slope = joistDto.Slope |> Option.defaultValue 0.0
        let pitchType = joistDto.PitchType |> Option.map PitchType.Parse
        let seatDepthLeft =
            joistDto.SeatDepthLeft
            |> Option.map
                (fun d ->
                    try
                        InchDotFraction(d).ToInch()
                    with
                    | _ -> failwith (sprintf "Mark %s's seat depth at the lef tis in an imporper format" takeoffMark) )
        let seatDepthRight =
            joistDto.SeatDepthRight
            |> Option.map
                (fun d ->
                    try
                        InchDotFraction(d).ToInch()
                    with
                    | _ -> failwith (sprintf "Mark %s's seat depth at the left is in an imporper format" takeoffMark) )
        
        let tcxlLength =
            joistDto.TcxlLength
            |> Option.map
                (fun l ->
                    try
                        FeetDotInch(l).ToFeet()
                    with
                    | _ -> failwith (sprintf "Mark %s's tcxl length is in an imporper format" takeoffMark) )

        let tcxlType = joistDto.TcxlType |> Option.map TcxType.Parse

        let tcxrLength =
            joistDto.TcxrLength
            |> Option.map
                (fun l ->
                    try
                        FeetDotInch(l).ToFeet()
                    with
                    | _ -> failwith (sprintf "Mark %s's tcxl length is in an imporper format" takeoffMark) )

        let tcxrType = joistDto.TcxrType |> Option.map TcxType.Parse
        let loads = [] |> List.toSeq
        let joist =
            let isOk result =
                match result with
                | Ok _ -> true
                | _ -> false
            
            let getError result =
                match result with
                | Ok _ -> ""
                | Error msg -> msg

            match (quantity |> isOk, designation |> isOk, baseLength |> isOk) with
            | true, true, true   ->         
                { Id = 0
                  TakeoffMark = takeoffMark
                  Quantity = quantity |> handleWithFailure
                  Designation = designation |> handleWithFailure
                  BaseLength = baseLength |> handleWithFailure
                  Slope = slope
                  PitchType = pitchType
                  SeatDepthLeft = seatDepthLeft
                  SeatDepthRight = seatDepthRight
                  TcxlLength = tcxlLength
                  TcxlType = tcxlType
                  TcxrLength = tcxrLength
                  TcxrType = tcxrType
                  Loads = loads
                } |> Ok
            | _ ->
                let errors =
                    seq { if not (isOk quantity) then yield getError quantity
                          if not (isOk designation) then yield getError designation
                          if not (isOk baseLength) then yield getError baseLength }
                Error errors
        joist
           
        
        
