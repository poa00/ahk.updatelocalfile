#SingleInstance, Force
#Noenv

; #include <oXMLManipulator>

sSettingFilePath := A_ScriptDir "\LocalFileUpdator.xml"

If ( ! FileExist( sSettingFilePath ) ) {
    Msgbox, 48,, The setting file could not be found. Program exists.
    ExitApp
}

oDOM := _getXMLCOMByFile( sSettingFilePath )

; Pattern 1
iCount := 0
for _nodeItem, iNodeNumber in oDOM.SelectNodes( "//Items/Item" ) {

    if ( !_nodeItem.childNodes.length ) {
        continue
    }
    _iInComment     := _nodeItem.selectSingleNode( "Version/InComment" ).text
    _sURL           := _nodeItem.selectSingleNode( "URL" ).text
    _sLocalPath     := _nodeItem.selectSingleNode( "LocalPath" ).text
    _sLocalPath     := getPathResolved( _sLocalPath, A_ScriptDir )
    _sLocalVersion  := getScriptVersion_FromFile( _sLocalPath, _iInComment, "0" )
    _sRemoteData    := getHeadingOfWebPage( _sURL, 100 )   ; first 100 bytes
    _sRemoteVersion := getScriptVersion( _sRemoteData, _iInComment, "0" )
    
    if ( -1 != compareVersions( _sLocalVersion, _sRemoteVersion ) ) {
        continue
    }
        
    ; If the local version is older than the remote one.
    bUpdated := updateFile( _sLocalPath, _sURL )
    if ( bUpdated ) {
        iCount++
    }
}

Msgbox % iCount++ " file(s) updated."
Return
updateFile( sLocalPath, sURL ) {

    if ( FileExist( sLocalPath ) ) {
        SplitPath, sLocalPath, name, _sDirPath, ext, _sNameWOExt, drive
        _sBackupPath := _sDirPath "\" _sNameWOExt ".bak"
        FileCopy, % sLocalPath, % _sBackupPath, 1   ; override existing file
        FileDelete, % sLocalPath
        bDeleted := ErrorLevel ? true : false
    }
    URLDownloadToFIle, % sURL, % sLocalPath
    bDownloadError := ErrorLevel
    if ( ErrorLevel && bDeleted ) {
        FileCopy, % _sBackupPath, % sLocalPath, 1   ; override existing file
    }
    FileDelete, % _sBackupPath
    return bDownloadError ? 0 : 1
    
}
_getXMLCOMByFile( sXMLFilePath ) {                        
    _oXML := _getXMLCOM()
    _oXML.load( sXMLFilePath )             
    _bsError := _isXMLParseError( _oXML )
    if ( _bsError ) {
        return _bsError
    }
    return _oXML
}
    _getXMLCOM() {
        static _sMSXML := "MSXML2.DOMDocument" (A_OSVersion ~= "WIN_(7|8)" ? ".6.0" : "")    
        _oXML := ComObjCreate( _sMSXML )
        _oXML.async := false
        _oXML.validateOnParse := false
        ; _oXML.resolveExternals := false       
        return _oXML
    } 
    /**
     * @return      boolean|string      false if no error; otherwise, the error message
     */
    _isXMLParseError( oXMLCOM ) {
        if ( oXMLCOM.parseError.errorCode != 0 ) {  
           return oXMLCOM.parseError
        }           
        return false
    }       
/**
 * Download top part of a specified Web page.
 * 
 * @see https://autohotkey.com/board/topic/69556-gui-cant-display-non-standard-characters/
 * @version     1.0.0
 */
getHeadingOfWebPage(URL, iMaxBytes=0, UserAgent = "", Proxy = "", ProxyBypass = "") {
    _sResult := ""
    ; Requires Windows Vista, Windows XP, Windows 2000 Professional, Windows NT Workstation 4.0,
    ; Windows Me, Windows 98, or Windows 95.
    ; Requires Internet Explorer 3.0 or later.
    pFix:=a_isunicode ? "W" : "A"

    hModule := DllCall("LoadLibrary", "Str", "wininet.dll") 

    AccessType := Proxy != "" ? 3 : 1
    ;INTERNET_OPEN_TYPE_PRECONFIG                    0   // use registry configuration 
    ;INTERNET_OPEN_TYPE_DIRECT                       1   // direct to net 
    ;INTERNET_OPEN_TYPE_PROXY                        3   // via named proxy 
    ;INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  4   // prevent using java/script/INS 

    io := DllCall("wininet\InternetOpen" . pFix
        , "Str", UserAgent ;lpszAgent 
        , "UInt", AccessType 
        , "Str", Proxy 
        , "Str", ProxyBypass 
        , "UInt", 0) ;dwFlags 

    iou := DllCall("wininet\InternetOpenUrl" . pFix
        , "UInt", io 
        , "Str", url 
        , "Str", "" ;lpszHeaders 
        , "UInt", 0 ;dwHeadersLength 
        , "UInt", 0x80000000 ;dwFlags: INTERNET_FLAG_RELOAD = 0x80000000 // retrieve the original item 
        , "UInt", 0) ;dwContext 

    If (ErrorLevel != 0 or iou = 0) { 
        DllCall( "FreeLibrary", "UInt", hModule ) 
        return 0 
    } 
    
    _iOffset := A_IsUnicode ? 2 : 0
    _iReservedVarCap := 16  ; 10240 
    VarSetCapacity( buffer, _iReservedVarCap, 0 )
    VarSetCapacity( BytesRead, 4, 0 )

    _iBytesReadTotal := 0
    Loop 
    { 
        ;http://msdn.microsoft.com/library/en-us/wininet/wininet/internetreadfile.asp
        irf := DllCall("wininet\InternetReadFile", "UInt", iou, "UInt", &buffer, "UInt", _iReservedVarCap, "UInt", &BytesRead) 
		VarSetCapacity(buffer, -1) ;to update the variable's internally-stored length

        BytesRead_ := 0 ; reset
		
        ; Build the integer by adding up its bytes. (From ExtractInteger-function)
        Loop, 4  { 
            ; Bytes read in this very DllCall 
            _iByte := ( A_Index - 1 )
            BytesRead_ += *( &BytesRead + _iByte ) << 8 * ( _iByte ) 
        }
        ; To ensure all data is retrieved, an application must continue to call the
        ; InternetReadFile function until the function returns TRUE and the lpdwNumberOfBytesRead parameter equals zero.
        If (irf = 1 and BytesRead_ = 0) {
            break
        }
        Else ; append the buffer's contents
        {
            a_isunicode ? buffer:=StrGet( &buffer, "CP0" )
            ; _sResult .= SubStr(buffer, 1, BytesRead_ * (a_isunicode ? 2 : 1))
            _sResult .= SubStr( buffer, 1, BytesRead_ * 1 )
        }
       
        ; optional: retrieve only a part of the file 
        if ( iMaxBytes > 0 ) {
            _iBytesReadTotal += BytesRead_
        }
        
	} Until ( _iBytesReadTotal > iMaxBytes ) ; only read the first x bytes will be a multiple of the buffer size, if the file is not smaller; trim if neccessary)

    DllCall("wininet\InternetCloseHandle",  "UInt", iou) 
    DllCall("wininet\InternetCloseHandle",  "UInt", io) 
    DllCall("FreeLibrary", "UInt", hModule)
    return _sResult
}

/**
 * Retrieves a verion number in a comment block with the `@version` annotation from a given file.
 * 
 * @version     1.1.0       Added the `bInComment` parameter.
 */
    /**
     * Retrieves a verion number in a comment block with the `@version` annotation from a given file.
     * 
     * @return     string       The found version number.
     */ 
    getScriptVersion( sScriptCode, bInComment=true, sDefaultVersion="0.0.1" ) {
        if ( bInComment ) {
            RegexMatch( sScriptCode, "O)/\*([^*]|[\r\n]|(\*+([^*/]|[\r\n])))*\*+/", oMatches )
            sScriptCode := oMatches[ 0 ]
        }
        RegexMatch( sScriptCode, "Omi)@version\s+\K.+?(?=(\s+)|$)", oMatches )
        if ( "" != oMatches[ 0 ] ) {
            return oMatches[ 0 ]
        }
        return sDefaultVersion  ; default    
    }
    /**
     * Retrives a version number in a comment block with the `@version` annotation from a file of the given path.
     *
     * @return      string      The found version number.
     */
    getScriptVersion_FromFile( sScriptPath, bInComment=true, sDefaultVersion="0.0.1" ) {
        FileRead, _sScriptCode, % sScriptPath
        return getScriptVersion( _sScriptCode, bInComment, sDefaultVersion )
    }

    
/**
 * Rosolves relative/short paths into long absolute.
 * 
 * ### Example
 * ```autohotkey
 * msgbox % getPathResolved( "..\test", A_ScriptDir )
 * msgbox % getPathResolved( "S:\project\functions\sample" )
 * ```
 * @version     1.0.0
 */
/**
 * @see         PathCombine : www.msdn.microsoft.com/en-us/library/bb773571(VS.85).aspx
 * @since       1.0.0
 */
getPathResolved( sPath, sWorkingDir="" ) {

    static _sAorW := A_IsUnicode ? "W" : "A"
    sWorkingDir   := sWorkingDir ? sWorkingDir : A_WorkingDir
    VarSetCapacity( _sAbsolutePath, 260, 0 )
    DllCall( "shlwapi\PathCombine" _sAorW, Str, _sAbsolutePath, Str, sWorkingDir, Str, sPath )
    
    ; Convert short paths into long.
    Loop, Files, % _sAbsolutePath, FD 
    {
        return % A_LoopFileLongPath
    }    
    return _sAbsolutePath
    
}

/**
 * Compares two versions such as `0.0.1` vs `0.0.2`.
 * 
 * Supports the following string notations (case-insensitive).
 *  - RC
 *  - Unstable
 *  - Beta 
 *  - b     : same as `Beta`
 *  - Alpha 
 *  - a     : same as `Alpha`
 *  - Dev
 *
 * ### Examples
 
 * ```autohotkey
 *  msgbox % compareVersions( "0.0.1", "0.0.2" ) ; -1
 *  msgbox % compareVersions( 10, 10.1 ) ; -1
 *  msgbox % compareVersions( "0.0.0.1", "0.0.0.0.2" ) ; 1
 *  msgbox % compareVersions( "0.0.1", "0.0.1b" ) ; 1
 *  msgbox % compareVersions( "1", "0.0.1b" ) ; 1
 *  msgbox % compareVersions( "0.0.1a", "0.0.1b" ) ; -1
 *  msgbox % compareVersions( "0.1", "0.1.0" ) ; 0
 * ```
 *
 * @requires    AutoHotkey v1.1.13 as it uses `StrSplit()`.
 * @return      integer      -1 when version A is older than B. 0 when version A is equal to version B. 1 when version A is newer than version B.
 * @version     0.0.1
 */
compareVersions( sVersionA, sVersionB ) {
    
    _aVersionA := StrSplit( sVersionA, "." )
    _aVersionB := StrSplit( sVersionB, "." )
    _iMaxIndex := compareVersions_getMaxNumberOfElements( _aVersionA, _aVersionB )    
    loop % _iMaxIndex {
        _iVersionA := compareVersions_getVersionSanitized( _aVersionA[ A_Index ] )
        _iVersionB := compareVersions_getVersionSanitized( _aVersionB[ A_Index ] )
        if ( _iVersionA > _iVersionB ) {
            return 1
        }
        if ( _iVersionA < _iVersionB ) {
            return -1
        }
    }
    return 0
    
}
    /**
     * Supports dev, RC, beta, b, alpha, a
     * @return      integer
     */
    compareVersions_getVersionSanitized( sVersion ) {
        sVersion := trim( sVersion )
        if ( "" = sVersion ) {
            return 0
        }
        _aLevels := { "RC": -100
            , "UNSTABLE": -200
            , "BETA": -300, "([^A-Za-z]|^)B([^A-Za-z]|$)": -300
            , "ALPHA": -400, "([^A-Za-z]|^)A([^A-Za-z]|$)": -400
            , "DEV": -500 }
        
        _iOverallCoefficient := 1
        for _sLevel, _iCoefficient in _aLevels {
            ; for exatct matches
            if ( RegexMatch( sVersion, "i)^" _sLevel "$" ) ) {
                return _iCoefficient
            }
            ; for partial matches such as `10b`, `3dev`
            if ( RegexMatch( sVersion, "i)" _sLevel ) ) {
                sVersion := RegexReplace( sVersion, "i)" _sLevel )
                if sVersion is integer
                {
                    return sVersion * _iCoefficient
                }
                ; Exceptional cases. Maybe mixed like `a20b`.
                _iOverallCoefficient := _iOverallCoefficient * _iCoefficient
            }
        }
        return sVersion * _iOverallCoefficient
        
    }
    /**
     * @return      integer     The found maximum index between the given two.
     */
    compareVersions_getMaxNumberOfElements( aA, aB ) {
        _aIndex := []
        _iAMax  := aA.MaxIndex() 
        _iBMax  := aB.MaxIndex() 
        _aIndex[ _iAMax ] := _iAMax
        _aIndex[ _iBMax ] := _iBMax
        return _aIndex.MaxIndex()
    }