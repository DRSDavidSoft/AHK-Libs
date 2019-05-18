/*
---------------------------------------------------------------------------
Function:
    To load configuration settings file.
---------------------------------------------------------------------------
*/

LoadSettings( _File = "Settings.ini", _Section = "", _Multi = false ) {
    Settings := []

    if ( _Section ) {
        SectionNames := _Section
    } else {
        IniRead, SectionNames, % _File
    }

    if ( StrSplit( SectionNames, "`n" ).MaxIndex() == 1 ) {
        _Section := SectionNames
    }

    Loop, Parse, SectionNames, `n
    {
        SectionName := A_LoopField

        if _Multi
            Settings[ SectionName ] := []
        else
            Prefix := ( _Section ? "" : SectionName . "_" )

        IniRead, SectionContent, % _File, % SectionName

        Loop, Parse, SectionContent, `n
        {
            KeyLine := InStr( A_LoopField, "=" )
            StringLeft, Variable, A_LoopField, KeyLine - 1
            StringRight, Value, A_LoopField, StrLen( A_LoopField ) - KeyLine
            Variable := RegExReplace( Variable, "\W", "_" )

            if _Multi
                Settings[ SectionName ][ Variable ] := Value
            else
                Settings[ Prefix . Variable ] := Value
        }
    }

    return Settings
}
