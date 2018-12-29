/*
---------------------------------------------------------------------------
Function:
    To load configuration settings file.
---------------------------------------------------------------------------
*/

LoadSettings( _File = "Settings.ini", _Section = "" ) {
    Settings := []

    if ( _Section ) {
        SectionNames := _Section
    } else {
        IniRead, SectionNames, % _File
    }

    Loop, Parse, SectionNames, `n
    {
        SectionName := A_LoopField
        IniRead, SectionContent, % _File, % SectionName

        Loop, Parse, SectionContent, `n
        {
            KeyLine := InStr( A_LoopField, "=" )
            StringLeft, Variable, A_LoopField, KeyLine - 1
            StringRight, Value, A_LoopField, StrLen( A_LoopField ) - KeyLine
            Variable := RegExReplace( Variable, "\W", "_" )
            Settings[ Variable ] := Value
        }
    }

    return Settings
}
