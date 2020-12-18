# +-----------------------------------------------------------------------------+ 
# | App.Name:        Add-URL2URLSET*.ps1                    | 
# | App.Description :    Adds a url to an existing URL Setin ISA or TMG        | 
# | Functionality:     Connects to a TMG server, looks for a URLSet indicated     | 
# |            as first parameter and attempts to add a URL indicated     | 
# |            as second Parameter.                    | 
# |                                        | 
# |      *** Must be run on a server with the Firewall Service running ***    | 
# |                                        | 
# +-----------------------------------------------------------------------------+ 
# | Current Version:    1.06                            | 
# |            By:    mcosentino                        | 
# |            Add-URL2URLSET at Marianok.com.ar            | 
# |            http://mscosentino-en.blogspot.com            | 
# |            http://mscosentino.blogspot.com                | 
# |                                        | 
# |          Date:    07/20/2010                        | 
# |                                        | 
# +- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -+ 
# 
# Must be run on a server with the Firewall Service running 
# 
 
    Param( 
            [string] $URLSet = $(throw "The Parameter -URLSet is required."), 
            [string] $NewURL = $(throw "The Parameter -NewURL is required.") 
        ) 
 
     
    $URLSet = $URLSet.trim() 
 
 
    # Let's clean up the NewURL to keep all entries standard 
 
    $MyNewURL = $NewURL.Trim() 
    $MyNewURL = $MyNewURL -replace("(^Http://)","") 
    $MyNewURL = $MyNewURL -replace("(^https://)","") 
    $MyNewURL = $MyNewURL -replace("(^www.)","") 
 
    if ($MyNewURL -match ("^.*\..*"))  
    { 
 
        # First Instantiate the ISA/TMG objects 
        try { 
            $root = new-object -comobject "FPC.Root" 
            $isaArray=$root.GetContainingArray() 
            write-host "Sucessfully Instantiated the TMG/ISA Objects. " 
        } 
        catch { 
            throw("Unable to Instantiated the TMG/ISA Objects. `n" +  $err.exception + $err) 
        } 
        finally { 
        } 
 
 
 
        # Second Get the URLSets 
        try { 
            $UrlSets=$isaArray.RuleElements.UrlSets 
        } 
        catch { 
            throw("Unable to retrieve List of URLSets from Array '" + $isaArray.name + "'`n" +  $err.exception + "`n" + $err) 
        } 
        finally { 
        } 
 
 
        # Third, Let's try to access the URLSet that we want to Update 
        try { 
            $MySet=$UrlSets.Item($URLSet) 
        } 
        catch { 
            throw("URLSet '$URLSet' Not found in Array '" + $isaArray.name + "'") 
        } 
        finally { 
        } 
 
 
     
        write-host "conected to URLSet:" $MySet.Name 
 
 
        # Fourth, Try to Add the URL 
        try { 
                $MySet.Add($MyNewURL) 
            $MySet.Save() 
            write-host $MyNewURL " `tAdded to the '$URLSet' URL Set"             
        } 
        catch { 
            $err=$Error[0] 
            # If Err.Number = -2147024713 
                if ($err.exception  -like "*file already exists*") { 
                write-host $MyNewURL " `twas Already in the '$URLSet' URL Set"             
                $err.clear 
            } 
            else { 
                $err.exception 
                $err 
            } 
        } 
        finally { 
        } 
 
     
 
        # Fifth, Try to add the rule for the New URL's sub domains 
        try { 
            $MySet.Add("*.$MyNewURL") 
            $MySet.Save() 
            write-host "*.$MyNewURL `tAdded to the '$URLSet' URL Set"             
        } 
        catch { 
            $err=$Error[0] 
            # If Err.Number = -2147024713 
                if ($err.exception  -like "*file already exists*") { 
                write-host "*.$MyNewURL `twas Already in the '$URLSet' URL Set"             
                $err.clear 
            } 
            else { 
                $err.exception 
                $err 
            } 
        } 
        finally { 
        } 
 
 
 
    } 
    else { 
        throw "The URL is invalid." 
    } 
 
 
     
    # Finally, list all URLs in the Set 
    write-host "" 
    write-host "URL Set '$URLSet' has the following entries:" 
    write-host "" 
    foreach ($url in $MySet) 
    { 
        write-host "URL Name:"  $url  
    }     
    write-host "" 
 
