$VerbosePreference = "Continue"

function OverpassGeoJson-ToMap
{
    param
    (
        [Parameter(ValueFromPipeline=$true)]$GeoJson,
        [String]$MapName = "GeoJson",
        $DefaultPointName = "New Point",
        $DefaultWayName = "New Way",
        [Switch]$OsmType
    )
    
    $Map = New-Object PSObject
    $Map | Add-Member -MemberType NoteProperty -Name "Name" -Value $MapName
    $Map | Add-Member -MemberType NoteProperty -Name "Version" -Value ([decimal]0.1.0)

    $folder = New-Object PSObject
    $folder | Add-Member -MemberType NoteProperty -Name "Name" -Value "Root"
    $folder | Add-Member -MemberType NoteProperty -Name "Type" -Value "Folder"
    $folder | Add-Member -MemberType NoteProperty -Name "Description" -Value ""
    $folder | Add-Member -MemberType NoteProperty -Name "Elements" -Value @()
    $folder | Add-Member -MemberType NoteProperty -Name "Subfolders" -Value @()

    $totalSteps = $Geojson.elements.count
    $currentStep = 0

    #Write-Progress -Activity "Converting Geojson to Map" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete $percentComplete

    foreach($element in $GeoJson.elements)
    {

        $placemark = New-Object PSObject
        $placemark | Add-Member -MemberType NoteProperty -Name "Name" -Value $DefaultPointName
        $placemark | Add-Member -MemberType NoteProperty -Name "Description" -Value ""
        $placemark | Add-Member -MemberType NoteProperty -Name "Properties" -Value @()
        $placemark | Add-Member -MemberType NoteProperty -Name "Coordinates" -Value @()

        $element.tags.PSObject.Properties | foreach {
            if($_.Name -notlike "_*")
            {
                $tag = New-Object PSObject
                $tag | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $tag | Add-Member -MemberType NoteProperty -Name "Value" -Value $_.Value
                $placemark.Properties += $tag
            }
        }

        if($OsmType)
        {
            $placemark.Name = $element.tags._osm_type
        }


        if($element.geometry.type -like "Point")
        {
            $placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Point"

            $coordinates = New-Object PSObject
            $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"-Value ([decimal]$element.geometry.coordinates[0])
            $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$element.geometry.coordinates[1])
            $placemark.Coordinates += $coordinates
        }

        if($element.geometry.type -like "LineString")
        {
            $placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Way"

            foreach($pair in $element.geometry.coordinates)
            {
                $coordinates = New-Object PSObject
                $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"-Value ([decimal]$pair[0])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$pair[1])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value 0
                $placemark.Coordinates += $coordinates
            }
        }

        if($placemark.Type) {
            $folder.Elements += $placemark    
        }
        

        $currentStep++
        $percentComplete = [int]($currentStep / $totalSteps * 100)
        #Write-Progress -Activity "Converting Geojson to Map" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete $percentComplete
    }

    #Write-Progress -Activity "Converting Geojson to Map" -Status "Completed!" -PercentComplete 100

    $Map | Add-Member -MemberType NoteProperty -Name "RootFolder" -Value $folder
    return $Map
}

function GeoJson-ToMap
{
    param
    (
        [Parameter(ValueFromPipeline=$true)]$GeoJson,
        [String]$MapName = "GeoJson",
        $DefaultFeatureName = "Feature"
    )
    
    $Map = New-Object PSObject
    $Map | Add-Member -MemberType NoteProperty -Name "Name" -Value $MapName
    $Map | Add-Member -MemberType NoteProperty -Name "Version" -Value ([decimal]0.1.0)

    $folder = New-Object PSObject
    $folder | Add-Member -MemberType NoteProperty -Name "Name" -Value "Root"
    $folder | Add-Member -MemberType NoteProperty -Name "Type" -Value "Folder"
    $folder | Add-Member -MemberType NoteProperty -Name "Description" -Value ""
    $folder | Add-Member -MemberType NoteProperty -Name "Elements" -Value @()
    $folder | Add-Member -MemberType NoteProperty -Name "Subfolders" -Value @()

    $totalSteps = $Geojson.features.count
    $percentComplete = 0
    $currentStep = 0

    #Write-Progress -Activity "Converting Geojson to Map" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete $percentComplete

    foreach($feature in $GeoJson.features)
    {

        $placemark = New-Object PSObject
        $placemark | Add-Member -MemberType NoteProperty -Name "Name" -Value $DefaultFeatureName
        $placemark | Add-Member -MemberType NoteProperty -Name "Description" -Value ""
        $placemark | Add-Member -MemberType NoteProperty -Name "Properties" -Value @()
        $placemark | Add-Member -MemberType NoteProperty -Name "Coordinates" -Value @()

        $feature.properties.PSObject.Properties | foreach {
            if($_.Name -notlike "_*")
            {
                $property = New-Object PSObject
                $property | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
                $property | Add-Member -MemberType NoteProperty -Name "Value" -Value $_.Value
                $placemark.Properties += $property
            }
        }


        if($placemark.Properties | where{$_.Name -like "name"}) {
            $placemark.Name = ($placemark.Properties | where{$_.Name -like "name"})[0].Value
        } else {
            if($placemark.Properties | where{$_.Name -like "address"}) {
                $placemark.Name = ($placemark.Properties | where{$_.Name -like "address"})[0].Value
            }
        }

        if($feature.geometry.type -like "Point")
        {
            $placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Point"

            $coordinates = New-Object PSObject
            $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"-Value ([decimal]$feature.geometry.coordinates[0])
            $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$feature.geometry.coordinates[1])
            $placemark.Coordinates += $coordinates
        }

        if($feature.geometry.type -like "LineString")
        {
            $placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Way"

            foreach($pair in $feature.geometry.coordinates)
            {
                $coordinates = New-Object PSObject
                $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"-Value ([decimal]$pair[0])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$pair[1])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value 0
                $placemark.Coordinates += $coordinates
            }
        }

        if($feature.geometry.type -like "Polygon")
        {

            $placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Polygon"

            #TODO - Currently only adding the outerline. Need second loop for inner lines.
            foreach($pair in $feature.geometry.coordinates[0]) {
                $coordinates = New-Object PSObject
                $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"-Value ([decimal]$pair[0])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$pair[1])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value 0
                $placemark.Coordinates += $coordinates
            }
        }

        if($placemark.Type) {
            $folder.Elements += $placemark    
        }
        

        $currentStep++
        $percentComplete = [int]($currentStep / $totalSteps * 100)
        #Write-Progress -Activity "Converting Geojson to Map" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete $percentComplete
    }

    #Write-Progress -Activity "Converting Geojson to Map" -Status "Completed!" -PercentComplete 100

    $Map | Add-Member -MemberType NoteProperty -Name "RootFolder" -Value $folder
    return $Map
}

function Kml-ToMap {
    param (
        [Parameter(ValueFromPipeline=$true)][System.Xml.XmlDocument]$Kml
    )


    if($Kml.kml.Document.name -eq $null) {
        Write-Error "Could not find kml.Document information. Invalid XML structure."
        return $null
    }

    if($Kml.kml.Document.Folder -eq $null) {
        Write-Error "Kml does not countain a root folder. Root folders are necessary for conversions."
        return $null
    }

    $Map = New-Object PSObject
    $Map | Add-Member -MemberType NoteProperty -Name "Name" -Value $Kml.kml.Document.name
    $Map | Add-Member -MemberType NoteProperty -Name "Version" -Value ([decimal]0.1.0)
    
    $Map | Add-Member -MemberType NoteProperty -Name "_KmlStyles" -Value @()
    $Map | Add-Member -MemberType NoteProperty -Name "_KmlStyleMap" -Value @()

    foreach($KmlStyle in $Kml.kml.Document.Style) {
        $Style = New-Object PSObject
        $Style | Add-Member -MemberType NoteProperty -Name "id" -Value $KmlStyle.id

        if($KmlStyle.IconStyle) {
            $IconStyle = New-Object PSObject
            $IconStyle | Add-Member -MemberType NoteProperty -Name "scale" -Value $KmlStyle.IconStyle.scale
        
            $HotSpot = New-Object PSOBject
            $HotSpot | Add-Member -MemberType NoteProperty -Name "X" -Value $KmlStyle.IconStyle.hotSpot.x
            $HotSpot | Add-Member -MemberType NoteProperty -Name "Y" -Value $KmlStyle.IconStyle.hotSpot.y
            $HotSpot | Add-Member -MemberType NoteProperty -Name "Xunits" -Value $KmlStyle.IconStyle.hotSpot.xunits
            $HotSpot | Add-Member -MemberType NoteProperty -Name "Yunits" -Value $KmlStyle.IconStyle.hotSpot.yunits
            $IconStyle | Add-Member -MemberType NoteProperty -Name "HotSpot" -Value $HotSpot
            $IconStyle | Add-Member -MemberType NoteProperty -Name "Icon" -Value $KmlStyle.IconStyle.Icon.href
            $Style | Add-Member -MemberType NoteProperty -Name "IconStyle" -Value $IconStyle
        }

        if($KmlStyle.LabelStyle)
        {
            $LabelStyle = New-Object PSOBject

            if($KmlStyle.LabelStyle.color) {
                $LabelStyle | Add-Member -MemberType NoteProperty -Name "Color" -Value $KmlStyle.LabelStyle.color
            }
            
            if($KmlStyle.LabelStyle.scale) {
                $LabelStyle | Add-Member -MemberType NoteProperty -Name "Scale" -Value ([decimal]$KmlStyle.LabelStyle.scale)
            }

            $Style | Add-Member -MemberType NoteProperty -Name "LabelStytle" -Value $LabelStyle
        }

        if($KmlStyle.LineStyle)
        {
            $LineStyle = New-Object PSOBject

            if($KmlStyle.LineStyle.color) {
                $LineStyle | Add-Member -MemberType NoteProperty -Name "Color" -Value $KmlStyle.LineStyle.color
            }
            
            if($KmlStyle.LineStyle.width) {
                $LineStyle | Add-Member -MemberType NoteProperty -Name "Width" -Value ([decimal]$KmlStyle.LineStyle.width)
            }

            $Style | Add-Member -MemberType NoteProperty -Name "LineStyle" -Value $LineStyle
        }

        if($KmlStyle.PolyStyle)
        {
            $PolyStyle = New-Object PSOBject

            if($KmlStyle.PolyStyle.color) {
                $PolyStyle | Add-Member -MemberType NoteProperty -Name "Color" -Value $KmlStyle.PolyStyle.color
            }

            $Style | Add-Member -MemberType NoteProperty -Name "PolyStyle" -Value $PolyStyle
        }
       
        if($KmlStyle.ListStyle)
        {
            $ListStyle = New-Object PSOBject

            if($KmlStyle.ListStyle.color) {
                $ListStyle | Add-Member -MemberType NoteProperty -Name "ItemIcon" -Value $KmlStyle.ListStyle.ItemIcon.href
            }

            $Style | Add-Member -MemberType NoteProperty -Name "ListStyle" -Value $ListStyle
        }

        $Map._KmlStyles += $Style

    }

    foreach($KmlStyleMap in $Kml.kml.Document.StyleMap) {
        $StyleMap = New-Object PSObject
        $StyleMap | Add-Member -MemberType NoteProperty -Name "id" -Value $KmlStyleMap.id
        $StyleMap | Add-Member -MemberType NoteProperty -Name "Urls" -value @()

        $KmlStyleMap.Pair | foreach {
            $SMUrl = New-Object PSObject
            $SMUrl | Add-Member -MemberType NoteProperty -Name "Key" -Value $_.key
            $SMUrl | Add-Member -MemberType NoteProperty -Name "StyleUrl" -Value $_.styleURL
            $StyleMap.Urls += $SMUrl
        }

        $Map._KmlStyleMap += $StyleMap
    }

    $folder = New-Object PSObject
    $folder | Add-Member -MemberType NoteProperty -Name "Name" -Value $Kml.kml.Document.Folder.name
    $folder | Add-Member -MemberType NoteProperty -Name "Type" -Value "Folder"
    $folder | Add-Member -MemberType NoteProperty -Name "Description" -Value $Kml.kml.Document.Folder.description
    $folder | Add-Member -MemberType NoteProperty -Name "Elements" -Value @()
    $folder | Add-Member -MemberType NoteProperty -Name "Subfolders" -Value @()

    $Map | Add-Member -MemberType NoteProperty -Name "RootFolder" -Value $folder


    if($Kml.kml.Document.Folder.Folder) {
      $Map.RootFolder.Subfolders += $Kml.kml.Document.Folder.Folder | foreach { $_ | KmlFolder-ToMap }
    }

    if($Kml.kml.Document.Folder.Placemark) {
       $Map.RootFolder.Elements += $Kml.kml.Document.Folder.Placemark | foreach { $_ | KmlPlacemark-ToMap }
    }

    return $Map
}

function KmlFolder-ToMap {
    param (
        [Parameter(ValueFromPipeline=$true)]$KmlFolder
    )

    $folder = New-Object PSObject
    $folder | Add-Member -MemberType NoteProperty -Name "Name" -Value $KmlFolder.name
    $folder | Add-Member -MemberType NoteProperty -Name "Type" -Value "Folder"
    $folder | Add-Member -MemberType NoteProperty -Name "Description" -Value $KmlFolder.description
    $folder | Add-Member -MemberType NoteProperty -Name "Elements" -Value @()
    $folder | Add-Member -MemberType NoteProperty -Name "Subfolders" -Value @()
    
    foreach($KmlSub in $KmlFolder.Folder) {
        $folder.Subfolders += $KmlSub | KmlFolder-ToMap
    }

    foreach($KmlPm in $KmlFolder.Placemark) {
        $folder.Elements += $KmlPm | KmlPlacemark-ToMap
    }

    return $folder
}

function KmlPlacemark-ToMap {
    param (
        [Parameter(ValueFromPipeline=$true)]$Placemark
    )

    $mapPlacemark = New-Object PSObject
    $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Name" -Value $Placemark.name
    $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Description" -Value $Placemark.description
    $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Properties" -Value @()
    $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Coordinates" -Value @()
    $mapPlacemark | Add-Member -MemberType NoteProperty -Name "_kmlStyleURL" -Value $Placemark.styleUrl

    if($Placemark.Point) {
        $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Point"
        
        $coordinates = New-Object PSObject
        $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude" -Value ([decimal]$Placemark.Point.coordinates.Split(",")[0])
        $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$Placemark.Point.coordinates.Split(",")[1])
        $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value ([decimal]$Placemark.Point.coordinates.Split(",")[2])

        $mapPlacemark.Coordinates += $coordinates
        
    }

    if($Placemark.Linestring) {
        $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Way"
        
        foreach($coord in $Placemark.LineString.coordinates.Replace("`n","").Replace("`t","").split(" ")) {
            if($coord -notlike "") {
                $coordinates = New-Object PSObject
                $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude" -Value ([decimal]$coord.Split(",")[0])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$coord.Split(",")[1])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value ([decimal]$coord.Split(",")[2])
                $mapPlacemark.Coordinates += $coordinates
            }
        }
    }


    if($Placemark.Polygon) {
        $mapPlacemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Polygon"

        #TODO - Implement inner line.
        foreach($coord in $Placemark.Polygon.outerBoundaryIs.LinearRing.coordinates.Replace("`n","").Replace("`t","").split(" ")) {
            if($coord -notlike "") {
                $coordinates = New-Object PSObject
                $coordinates | Add-Member -MemberType NoteProperty -Name "Latitude" -Value ([decimal]$coord.Split(",")[0])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Longitude" -Value ([decimal]$coord.Split(",")[1])
                $coordinates | Add-Member -MemberType NoteProperty -Name "Altitude" -Value ([decimal]$coord.Split(",")[2])
                $mapPlacemark.Coordinates += $coordinates
            }
        }
    }

    return $mapPlacemark
}

function Map-ToKml {
    param
    (
        [Parameter(ValueFromPipeline=$true)]$Map
    )
    
    $KmlStd = "http://www.opengis.net/kml/2.2"

    $Kml = New-Object System.Xml.XmlDocument
    $RootKml = $Kml.CreateElement("kml",$KmlStd)
    $Kml.AppendChild($RootKml) | Out-Null

    $Document = $Kml.CreateElement("Document",$KmlStd)
    
    $DocumentName = $Kml.CreateElement("name",$KmlStd)
    $DocumentName.InnerText = $Map.Name

    $Document.AppendChild($DocumentName) | Out-Null
 
    foreach($MapStyleLink in $Map._KmlStyleMap) {
        $StyleMap = $Kml.CreateElement("StyleMap",$KmlStd)
        $StyleMap.SetAttribute("id",$MapStyleLink.id)
        
        foreach($MapPair in $MapStyleLink.Urls)
        {
            $Pair = $Kml.CreateElement("Pair",$KmlStd)
            
            $key = $Kml.CreateElement("key",$KmlStd)
            $key.InnerText = $MapPair.Key
            $Pair.AppendChild($key) | Out-Null

            $styleURL = $Kml.CreateElement("styleUrl",$KmlStd)
            $styleURL.InnerText = $MapPair.StyleUrl
            $Pair.AppendChild($styleURL) | Out-Null

            $StyleMap.AppendChild($Pair) | Out-Null
        }
        $Document.AppendChild($StyleMap) | Out-Null
        
    }   

    foreach($MapStyle in $Map._KmlStyles) {
        $Style = $Kml.CreateElement("Style",$KmlStd)
        $Style.SetAttribute("id",$MapStyle.id)

        if($MapStyle.IconStyle) {
            $IconStyle = $Kml.CreateElement("IconStyle",$KmlStd)
                if($MapStyle.IconStyle.Scale) {
                    $Scale = $Kml.CreateElement("scale",$KmlStd)
                    $Scale.InnerText = $MapStyle.IconStyle.Scale
                    $IconStyle.AppendChild($Scale) | Out-Null
                }
                if($MapStyle.IconStyle.Icon) {
                    $Icon = $Kml.CreateElement("Icon",$KmlStd)
                    $href = $Kml.CreateElement("href",$KmlStd)
                    $href.InnerText = $MapStyle.IconStyle.Icon
                    $Icon.AppendChild($href) | Out-Null
                    $IconStyle.AppendChild($Icon) | Out-Null
                }
                if($MapStyle.IconStyle.HotSpot) {
                    $HotSpot = $Kml.CreateElement("hotSpot",$KmlStd)
                    $HotSpot.SetAttribute("x",$MapStyle.IconStyle.HotSpot.X)
                    $HotSpot.SetAttribute("y",$MapStyle.IconStyle.HotSpot.Y)
                    $HotSpot.SetAttribute("xunits",$MapStyle.IconStyle.HotSpot.Xunits)
                    $HotSpot.SetAttribute("yunits",$MapStyle.IconStyle.HotSpot.Yunits)
                    $IconStyle.AppendChild($HotSpot) | Out-Null
                }
            $Style.AppendChild($IconStyle) | Out-Null
        }

        if($MapStyle.LabelStyle) {
            $LabelStyle = $Kml.CreateElement("LabelStyle",$KmlStd)
            if($MapStyle.LabelStyle.color) {
                $color = $Kml.CreateElement("color",$KmlStd)
                $color.InnerText = $MapStyle.LabelStyle.Color
                $LabelStyle.AppendChild($color) | Out-Null
            }
            if($MapStyle.LabelStyle.scale) {
                $scale = $Kml.CreateElement("scale",$KmlStd)
                $scale.InnerText = $MapStyle.LabelStyle.Scale
                $LabelStyle.AppendChild($scale) | Out-Null
            }
            $Style.AppendChild($LabelStyle) | Out-Null
        }

        if($MapStyle.LineStyle) {
            $LineStyle = $Kml.CreateElement("LineStyle",$KmlStd)
            if($MapStyle.LineStyle.color) {
                $color = $Kml.CreateElement("color",$KmlStd)
                $color.InnerText = $MapStyle.LineStyle.Color
                $LineStyle.AppendChild($color) | Out-Null
            }
            if($MapStyle.LineStyle.Width) {
                $Width = $Kml.CreateElement("width",$KmlStd)
                $Width.InnerText = $MapStyle.LineStyle.Width
                $LineStyle.AppendChild($Width) | Out-Null
            }
            $Style.AppendChild($LineStyle) | Out-Null
        }

        if($MapStyle.PolyStyle) {
            $PolyStyle = $Kml.CreateElement("PolyStyle",$KmlStd)
            if($MapStyle.PolyStyle.color) {
                $color = $Kml.CreateElement("color",$KmlStd)
                $color.InnerText = $MapStyle.PolyStyle.Color
                $PolyStyle.AppendChild($color) | Out-Null
            }
            $Style.AppendChild($PolyStyle) | Out-Null
        }

        if($MapStyle.ListStyle) {
            $ListStyle = $Kml.CreateElement("ListStyle",$KmlStd)
            if($MapStyle.ListStyle.ItemIcon) {
                $ItemIcon = $Kml.CreateElement("ItemIcon",$KmlStd)
                $href = $Kml.CreateElement("href",$KmlStd)
                $href.InnerText = $MapStyle.ListStyle.ItemIcon
                $ItemIcon.AppendChild($href) | Out-Null
                $ListStyle.AppendChild($ItemIcon) | Out-Null
            }
            $Style.AppendChild($ListStyle) | Out-Null
        }

        $Document.AppendChild($Style) | Out-Null
    }

    
    $RootKml.AppendChild($Document) | Out-Null

    $Folder = $Map.RootFolder | MapFolder-ToKml -Kml $Kml
    
    $Document.AppendChild($Folder) | Out-Null
    
    return $Kml
    
}

function MapFolder-ToKml {
    param
    (
        [Parameter(ValueFromPipeline=$true)]$Folder,
        $Kml,
        $KmlStd = "http://www.opengis.net/kml/2.2",
        [switch]$BulletPointComments
    )

    $KmlFolder = $Kml.CreateElement("Folder",$KmlStd)
    $KmlFolderName = $Kml.CreateElement("name",$KmlStd)
    $KmlFolderName.InnerText = $Folder.Name

    $KmlFolderDescription = $Kml.CreateElement("description",$KmlStd)
    $KmlFolderDescription.InnerText = $Folder.Description

    $KmlFolder.AppendChild($KmlFolderName) | Out-Null
    $KmlFolder.AppendChild($KmlFolderDescription) | Out-Null

    $totalSteps = $Folder.Elements.count
    $currentStep = 0

    #Write-Progress -Activity "Converting $($Folder.name) to Kml" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete 0
    
    foreach($element in $Folder.Elements)
    {
        $Placemark = $Kml.CreateElement("Placemark",$KmlStd)

        $PlacemarkName = $Kml.CreateElement("name",$KmlStd)
        
        $PlacemarkName.InnerText = $element.Name

        $PlacemarkDescription = $Kml.CreateElement("description",$KmlStd)
        
        $PlacemarkDescription.InnerText = $element.Description
        
        #TODO Use same function for table creation

        if($element.Properties.Count -gt 0)
        {
            $PlacemarkDescription.InnerText += "<br><br>`n"
            $PlacemarkDescription.InnerText += "<table style=`"width:250px; padding-bottom:10px;`">`n"
            $element.Properties| foreach {
                $PlacemarkDescription.InnerText += "<tr>`n"
                $PlacemarkDescription.InnerText += "<td><b>`n"
                $PlacemarkDescription.InnerText += $_.Name
                $PlacemarkDescription.InnerText += "</b></td>`n"
                $PlacemarkDescription.InnerText += "<td>`n"
                $PlacemarkDescription.InnerText += $_.Value
                $PlacemarkDescription.InnerText += "</td>`n"
                $PlacemarkDescription.InnerText += "</tr>`n"
            }
            $PlacemarkDescription.InnerText += "</table>"
        }

        $Placemark.AppendChild($PlacemarkName) | Out-Null
        $Placemark.AppendChild($PlacemarkDescription) | Out-Null
        
        if($element._kmlStyleURL -ne $null) {
            $StyleURL = $Kml.CreateElement("styleUrl",$KmlStd)
            $StyleURL.InnerText = $element._kmlStyleURL
            $Placemark.AppendChild($StyleURL) | Out-Null
        }



        if($element.Type -like "Way")
        {
            $PlacemarkLinestring = $Kml.CreateElement("LineString",$KmlStd)
        
            $PlacemarkTessellate = $Kml.CreateElement("tesselate",$KmlStd)
            $PlacemarkTessellate.InnerText = "1"

            $PlacemarkLinestringCoordinates = $Kml.CreateElement("coordinates",$KmlStd)
            $element.Coordinates | foreach {
                $PlacemarkLinestringCoordinates.InnerText += " " + $_.Latitude.ToString() + ","
                $PlacemarkLinestringCoordinates.InnerText += $_.Longitude.ToString() + ","
                $PlacemarkLinestringCoordinates.InnerText += $_.Altitude.ToString()
            }
            $PlacemarkLinestringCoordinates.InnerText = $PlacemarkLinestringCoordinates.InnerText.Substring(1)

            $PlacemarkLinestring.AppendChild($PlacemarkTessellate) | Out-Null
            $PlacemarkLinestring.AppendChild($PlacemarkLinestringCoordinates) | Out-Null
            $Placemark.AppendChild($PlacemarkLinestring) | Out-Null
        }

        if($element.Type -like "Polygon")
        {

            $PlacemarkPolygon = $Kml.CreateElement("Polygon",$KmlStd)

            $PlacemarkOuterBoundary = $Kml.CreateElement("outerBoundaryIs",$KmlStd)

            $PlacemarkLinearRing = $Kml.CreateElement("LinearRing",$KmlStd)
        
            $PlacemarkTessellate = $Kml.CreateElement("tesselate",$KmlStd)
            $PlacemarkTessellate.InnerText = "1"

            $PlacemarkLinestringCoordinates = $Kml.CreateElement("coordinates",$KmlStd)
            $element.Coordinates | foreach {
                $PlacemarkLinestringCoordinates.InnerText += " " + $_.Latitude.ToString() + ","
                $PlacemarkLinestringCoordinates.InnerText += $_.Longitude.ToString() + ","
                $PlacemarkLinestringCoordinates.InnerText += $_.Altitude.ToString()
            }
            $PlacemarkLinestringCoordinates.InnerText = $PlacemarkLinestringCoordinates.InnerText.Substring(1)

            $PlacemarkPolygon.AppendChild($PlacemarkTessellate) | Out-Null

            $PlacemarkLinearRing.AppendChild($PlacemarkLinestringCoordinates) | Out-Null
            $PlacemarkOuterBoundary.AppendChild($PlacemarkLinearRing) | Out-Null
            $PlacemarkPolygon.AppendChild($PlacemarkOuterBoundary) | Out-Null

            $Placemark.AppendChild($PlacemarkPolygon) | Out-Null
        }

        if($element.Type -like "Point")
        {
            $PlacemarkPoint = $Kml.CreateElement("Point",$KmlStd)

            $PlacemarkPointCoordinates = $Kml.CreateElement("coordinates",$KmlStd)
            $PlacemarkPointCoordinates.InnerText += " " + $element.Coordinates[0].Latitude.ToString() + ","
            $PlacemarkPointCoordinates.InnerText += $element.Coordinates[0].Longitude.ToString() + ","
            $PlacemarkPointCoordinates.InnerText += $element.Coordinates[0].Altitude.ToString()
            $PlacemarkPointCoordinates.InnerText = $PlacemarkPointCoordinates.InnerText.Substring(1)

            $PlacemarkPoint.AppendChild($PlacemarkPointCoordinates) | Out-Null
            $Placemark.AppendChild($PlacemarkPoint) | Out-Null
        }

        $KmlFolder.AppendChild($Placemark) | Out-Null

        $currentStep++
        $percentComplete = [int]($currentStep / $totalSteps * 100)
        ##Write-Progress -Activity "Converting $($Folder.name) to Kml" -Status "Processing Step $currentStep of $totalSteps" -PercentComplete $percentComplete
    }

    ##Write-Progress -Activity "Converting $($Folder.name) to Kml" -Status "Completed!" -PercentComplete 100

    $Folder.Subfolders | foreach {
        $Subfolder = $_ | MapFolder-ToKml -Kml $Kml
        $KmlFolder.AppendChild($Subfolder) | Out-Null
    }

    return $KmlFolder
}

function Read-Excel {
    param (
        [String]$Path
    )

    try {
        $ExcelPath = Get-ChildItem $Path | Select -ExpandProperty FullName
        $Excel = New-Object -ComObject Excel.Application
        
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        $tries = 0
        $timeout = 200 

        While ($Excel.Interactive -eq $true -or $tries -eq $timeout) {
            Write-Verbose ("Excel: Application unlocked. Trying to lock: " + $tries)
            try {
                $Excel.Interactive = $false;
            } catch {
                Write-Verbose $_
            }
        }

        $WB = $Excel.Workbooks.Open($ExcelPath)
        sleep -Milliseconds 800

        $WB.Worksheets[1].Activate()

        $table = @()

        $col = 1
        $Columns = @()
        While($WB.ActiveSheet.Cells.Item(1,$col).Value() -ne $Null) {
            $Columns += [String]($WB.ActiveSheet.Cells.Item(1,$col).Text)
            $col += 1
        }

        $row = 2
        While($WB.ActiveSheet.Cells.Item($row,1).Value() -ne $Null) {
            $data = New-Object PSObject
            for($col = 1;$col -le $Columns.Count;$col++) {
                $data | Add-Member -MemberType NoteProperty -Name $Columns[$col-1] -Value $WB.ActiveSheet.Cells.Item($row,$col).Text
            }
            $table += $data
            $row += 1
        }
        Write-Verbose "Excel: Success reading. Closing COM application."
        sleep -Milliseconds 800
        $Excel.Quit()
    } Catch {
        Write-Verbose ("Excel: Error reading file:" + $_ + "Closing COM application.")
        $Excel.Quit()
        Throw $_
    } Finally {
        Write-Verbose "Excel: Finalizing. Closing COM application."
        $Excel.Quit()
    }
    Return $table
}

function Count-MapFolderPlacemarks {
    param (
        [Parameter(ValueFromPipeline=$true)]$Folder,
        $Type
    )

    $NumberOfPlacemarks = 0

    if($Type) {
        $NumberOfPlacemarks = $Folder.Elements | Where {$_.Type -like $Type} | Measure | Select -ExpandProperty Count
    } else {
        $NumberOfPlacemarks = $Folder.Elements.Count  
    }
    
    foreach($Sub in $Folder.Subfolders) {
        $NumberOfPlacemarks += $Sub | Count-MapFolderPlacemarks -Type:$Type
    }

    return $NumberOfPlacemarks
}

function Count-MapPlacemarks {
    param (
        [Parameter(ValueFromPipeline=$true)]$Map,
        $Type
    )
    return $Map.RootFolder | Count-MapFolderPlacemarks -Type:$Type
}

function Extract-MapFolderPlacemarks {
    param (
        [Parameter(ValueFromPipeline=$true)]$Folder,
        $Type
    )
    $Placemarks = @()

    if($Type) {
        $Placemarks += $Folder.Elements | Where {$_.Type -like $Type}
    } else {
        $Placemarks += $Folder.Elements
    }
    
    foreach($Sub in $Folder.Subfolders) {
        $Placemarks += $Sub | Extract-MapFolderPlacemarks -Type:$Type
    }

    return $Placemarks
}


function Extract-MapPlacemarks {
    param (
        [Parameter(ValueFromPipeline=$true)]$Map,
        $Type
    )
    return $Map.RootFolder | Extract-MapFolderPlacemarks -Type:$Type
}



function SearchReplace-MapPlacemark {
    param (
        [Parameter(ValueFromPipeline=$true)]$Folder,
        $Data,
        [string[]]$BulletPointFields,
        [string]$NameField
    )
    process {
        $affected = 0
        #TODO Isolate table creation function

        if(!$NameField) {
            if($Data.Name) {
                $NameField = "Name"
            } else {
                if($Data.Address) {
                    $NameField = "Address"
                } else {
                    throw "Couldn't find a naming field to distinguished the data. Either add a 'Name' or 'Address' column or specify a different column with the '-NameField' argument."
                }
            }
        }

        foreach($Element in $Folder.Elements) {
            if($Element.Name -like $Data."$NameField") {
                $Element.Description = "<table style=`"width:400px; padding-bottom:10px;border: 1px solid #D3D3D3; border-collapse: collapse;`">`n"
                foreach($prop in $Data.PSObject.Properties) {
                    $Element.Description += "<tr style=`"width:400px;border: 1px solid #D3D3D3; border-collapse: collapse;`">`n"
                    $Element.Description += "`t<td style=`"text-align:left;width:150px;vertical-align:top;`"><b>"
                    $Element.Description += $prop.Name
                    $Element.Description += "</b></td>`n"
                    $Element.Description += "`t<td style=`"text-align:left;vertical-align:top;`" >"
                    if($BulletPointFields) {
                        if($BulletPointFields.Contains($prop.Name) -and $prop.Value) {
                            #$Element.Description += "<ul>"
                            #$prop.Value.split("`n") | foreach {$Element.Description += "<li>"+ $_ +"</li>"}
                            #$Element.Description += "</ul>"
                            $prop.Value.split("`n") | foreach {$Element.Description += " - "+ $_ +"<br>"}
                        } else {
                            $Element.Description += $prop.Value
                        }
                    } else {
                        $Element.Description += $prop.Value
                    }

                    $Element.Description += "`t</td>`n"
                    $Element.Description += "</tr>`n"
                }
                $Element.Description += "</table>"
                $affected += 1
            }
        }
    
        foreach($Sub in $Folder.Subfolders) {
            $affected += $Sub | SearchReplace-MapPlacemark -Data $Data -NameField:$NameField -BulletPointFields:$BulletPointFields
        }

        return $affected;
    }
}

function Map-ToGeojJson {
    param
    (
        [Parameter(ValueFromPipeline=$true)]$Map
    )
    
    $GeoJson = New-Object PSObject
    $GeoJson | Add-Member -MemberType NoteProperty -Name "type" -Value "FeatureCollection"
    $GeoJson | Add-Member -MemberType NoteProperty -Name "features" -Value @()
    $GeoJson.features += $Map.RootFolder | MapFolder-ToGeoJson
}

function MapFolder-ToGeoJson {
    param
    (
        [Parameter(ValueFromPipeline=$true)]$Folder
    )

    $features = @()
    
}

<#
$Map = New-Object PSObject
$Map | Add-Member -MemberType NoteProperty -Name "Name" -Value $MapName
$Map | Add-Member -MemberType NoteProperty -Name "Version" -Value ([decimal]0.1.0)
$Map | Add-Member -MemberType NoteProperty -Name "RootFolder" -Value $folder

$folder = New-Object PSObject
$folder | Add-Member -MemberType NoteProperty -Name "Name" -Value $FolderName
$folder | Add-Member -MemberType NoteProperty -Name "Type" -Value "Folder"
$folder | Add-Member -MemberType NoteProperty -Name "Description" -Value $FolderDescription
$folder | Add-Member -MemberType NoteProperty -Name "Elements" -Value @()
$folder | Add-Member -MemberType NoteProperty -Name "Subfolders" -Value @()

$placemark = New-Object PSObject
$placemark | Add-Member -MemberType NoteProperty -Name "Name" -Value "New Placemark"
$placemark | Add-Member -MemberType NoteProperty -Name "Type" -Value "Point" (Way | Point)
$placemark | Add-Member -MemberType NoteProperty -Name "Description" -Value ""
$placemark | Add-Member -MemberType NoteProperty -Name "Properties" -Value @{}
$placemark | Add-Member -MemberType NoteProperty -Name "Coordinates" -Value @()

$coordinates = New-Object PSObject
$coordinates | Add-Member -MemberType NoteProperty -Name "Longitude"
$coordinates | Add-Member -MemberType NoteProperty -Name "Latitude"
$coordinates | Add-Member -MemberType NoteProperty -Name "Altitude"
#>

<#
$Geojson = Get-Content '.\British Columbia Grid.json' -Encoding UTF8 | ConvertFrom-Json
$MQGrid = $Geojson | GeoJson-ToMap -MapName "British Columbia PowerGrid"
$MQGrid.RootFolder.Elements = $MQGrid.RootFolder.Elements | where{ ($_.Properties | where {$_.Name -like "power"}).Value -notlike "minor_line" }
$QGrid = $MQGrid | Map-ToKml
$QGrid.OuterXml | Out-File "British Columbia Grid.kml" -Encoding utf8
#>

<#
$Kml = [System.Xml.XmlDocument](Get-Content .\SimpleTest.kml)
$Map = $Kml | Kml-ToMap
$KmlRemake = $Map | Map-ToKml
$KmlRemake.OuterXml | Out-File "SimpleTest_remake.kml" -Encoding utf8
#>


function Prompt-FileDialog {
    param (
        $InitialDirectory,
        $FileFilter
    )
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    if($InitialDirectory) {
        $fileDialog.InitialDirectory = $InitialDirectory
    } else {
        $fileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    }

    if($FileFilter) {
        $fileDialog.Filter = "$FileFilter|$FileFilter"
    }

    if($fileDialog.ShowDialog() -eq "OK") {
        return $fileDialog.FileName
    }
}

<#

$LoKml_fileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "Google Earth KML (*.kml)|*.kml"
}

if($LoKml_fileDialog.ShowDialog() -eq "OK") {
    $KmlData = [System.Xml.XmlDocument](Get-Content $LoKml_FileDialog.FileName) 
    $MapData = $KmlData | Kml-ToMap
}




$ExcelPath = ".\SimpletTest.xlsx"
$KmlPath = ".\SimpleTest.kml"


$Kml = [System.Xml.XmlDocument](Get-Content $KmlPath)
$Info = Read-Excel $ExcelPath
$MapData = $Kml | Kml-ToMap

foreach($data in $ExcelData){
    $affected = $MapData.RootFolder | SearchReplace-MapPlacemark -Data $data -BulletPointField "Comments"
    Write-Host ($data.Name + ": " + $affected + " affected.")
}

$KmlRemake = $MapData | Map-ToKml
$KmlRemake.OuterXml | Out-File "SimpleTest_remake.kml" -Encoding utf8

#>