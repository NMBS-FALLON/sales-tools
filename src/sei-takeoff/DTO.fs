module Design.SalesTools.Sei.Dto  

type JoistDto =
    { Mark : string option
      Quantity : int option
      Depth : float option
      Series : string option
      Designation : string option
      BaseLength : string option
      Slope : float option
      PitchType : string option
      SeatDepthLeft : string option
      TcxlLength : string option
      TcxlType : string option
      TcxrLength : string option
      TcxrType : string option
      SeatDepthRight : string option
      AddLoad : float option
      NetUplift : float option
      AxialLoad : float option
      AxialType : string option
      Notes : string option }

    member joistDto.IsEmpty =
      joistDto.Mark.IsNone &&
      joistDto.Quantity.IsNone &&
      joistDto.Series.IsNone &&
      joistDto.Designation.IsNone &&
      joistDto.BaseLength.IsNone




