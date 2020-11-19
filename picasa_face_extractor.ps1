# Scan for .picasa.ini files and convert them 'faces' entries into XMP metadata
#
# The XMP metadata is modelled on the metadata produces by DigiKam 7.4, which
# includes a number of tags for various software packages.

#using assembly System.Xml.Linq # Doesn't work? Use Add-Type instead
using namespace System.Xml.Linq 

Import-Module PsIni
Add-Type -AssemblyName System.Xml.Linq 

$contactSecion = "Contacts2"

#$ErrorActionPreference = "Inquire"

#region picasa.ini

function Test-PicasaIniContainsContacts([System.Collections.Specialized.OrderedDictionary]$iniFile) {
    $iniFile.Contains($contactSecion)
}

# See https://gist.github.com/fbuchinger/1073823
function ConvertTo-TypedCropRectangle([string]$cropRectangle) {
    [PSCustomObject]@{
        Left = [System.Convert]::ToInt32($cropRectangle.Substring(0,4), 16)/65536
        Top = [System.Convert]::ToInt32($cropRectangle.Substring(4,4), 16)/65536
        Right = [System.Convert]::ToInt32($cropRectangle.Substring(8,4), 16)/65536
        Bottom = [System.Convert]::ToInt32($cropRectangle.Substring(12), 16)/65536
    }
}

# XML area is x,y,w,h, with x,y centered in the middle of the rectangle.
# Picasa rectangle is left,top,right,bottom.
function ConvertTo-XmpRectangle($cropRectangle) {
    [PSCustomObject]@{
        x = $cropRectangle.Left + ($cropRectangle.Right-$cropRectangle.Left)/2
        y = $cropRectangle.Top + ($cropRectangle.Bottom-$cropRectangle.Top)/2
        w = $cropRectangle.Right-$cropRectangle.Left
        h = $cropRectangle.Bottom-$cropRectangle.Top
    }
}

#endregion picasa.ini

#region Xmp

$nsRDF = [XNamespace]"http://www.w3.org/1999/02/22-rdf-syntax-ns#"
$nsMWG = [XNamespace]"http://www.metadataworkinggroup.com/schemas/regions/"
$nsStArea = [XNamespace]"http://ns.adobe.com/xmp/sType/Area#"
$nsDC = [XNamespace]"http://purl.org/dc/elements/1.1/"
$nsDigiKam = [XNamespace]"http://www.digikam.org/ns/1.0/"
$nsMicrosoftPhoto = [XNamespace]"http://ns.microsoft.com/photo/1.0/"
$nsLightRoom = [XNamespace]"http://ns.adobe.com/lightroom/1.0/"
$nsMediaPro = [XNamespace]"http://ns.iview-multimedia.com/mediapro/1.0/"
$nsAdobe = [XNamespace]"adobe:ns:meta/"

function ConvertTo-RdfLi($nodeContents) {
    $nodeContents | ForEach-Object { [XElement]::new($nsRDF + "li", $_.ToString()) }
}

function ConvertTo-RdfBag($nodeContents) {
    [XElement]::new($nsRDF + "Bag", (ConvertTo-RdfLi $nodeContents))
}

function ConvertTo-RdfSeq($nodeContents) {
    [XElement]::new($nsRDF + "Seq", (ConvertTo-RdfLi $nodeContents))
}

function ConvertTo-XmpArea($xmpRectangle) {
    [XElement]::new($nsMWG + "Area", @(
        [XAttribute]::new($nsStArea + 'x', $xmpRectangle.x),
        [XAttribute]::new($nsStArea + 'y', $xmpRectangle.y),
        [XAttribute]::new($nsStArea + 'w', $xmpRectangle.w),
        [XAttribute]::new($nsStArea + 'h', $xmpRectangle.h)
    ))
}

function ConvertTo-RdfFaceDescription($picasaFace) {
    [XElement]::new($nsRDF + "li", @(
        [XElement]::new($nsRDF + "Description", @(
            [XAttribute]::new($nsMWG + 'Name', $picasaFace.Contact),
            [XAttribute]::new($nsMWG + 'Type', 'Face'),
            (ConvertTo-XmpArea (ConvertTo-XmpRectangle (ConvertTo-TypedCropRectangle $picasaFace.Rectangle)))
        ))
    ))
}

function ConvertTo-XmpRegions($people) {
    [XElement]::new($nsMWG + "Regions", @(
        [XAttribute]::new($nsRDF + 'parseType', "Resource")
        [XElement]::new($nsMWG + "RegionList", @(
            [XElement]::new($nsRDF + "Bag", ($people | ForEach-Object { ConvertTo-RdfFaceDescription $_ }))
        ))
    ))
}

function ConvertTo-Xmp($people) {
    [XDocument]::new( 
        [XDeclaration]::new('1.0', 'utf-8', 'yes'), 
        [XElement]::new($nsAdobe + 'xmpmeta', @( 

            [XAttribute]::new([XNamespace]::Xmlns + "x", $nsAdobe), 
            [XAttribute]::new([XNamespace]::Xmlns + "rdf", $nsRDF), 
            [XAttribute]::new([XNamespace]::Xmlns + "stArea", $nsStArea), 
            [XAttribute]::new([XNamespace]::Xmlns + "mwg-rs", $nsMWG), 
            [XAttribute]::new([XNamespace]::Xmlns + "dc", $nsDC), 
            [XAttribute]::new([XNamespace]::Xmlns + "digikam", $nsDigiKam), 
            [XAttribute]::new([XNamespace]::Xmlns + "microsoftPhoto", $nsMicrosoftPhoto), 
            [XAttribute]::new([XNamespace]::Xmlns + "lr", $nsLightRoom), 
            [XAttribute]::new([XNamespace]::Xmlns + "mediapro", $nsMediaPro), 

            [XAttribute]::new($nsAdobe + 'xmptk', "XMP Core 4.4.0-Exiv2"),

            [XElement]::new($nsRDF + "RDF",
                [XElement]::new($nsRDF + "Description", @(
                    [XElement]::new($nsDC + "subject", (ConvertTo-RdfBag  ($people | ForEach-Object { $_.Contact }))),
                    [XElement]::new($nsDigiKam + "TagsList", (ConvertTo-RdfSeq ($people | ForEach-Object { $_.Contact }))),
                    [XElement]::new($nsMicrosoftPhoto + "LastKeywordXMP", (ConvertTo-RdfBag  ($people | ForEach-Object { $_.Contact }))),
                    [XElement]::new($nsLightRoom + "hierarchicalSubject", (ConvertTo-RdfBag  ($people | ForEach-Object { $_.Contact }))),
                    [XElement]::new($nsMediaPro + "CatalogSets", (ConvertTo-RdfBag  ($people | ForEach-Object { $_.Contact }))),
                    (ConvertTo-XmpRegions $people)
                ))
            )
        ))
    )
}

#endregion XMP


$folder = "\\Desktop-27m6eeq\D\Users\Bancrofts\pictures\old mo pics"

$files = Get-ChildItem -Path $folder -Include .picasa.ini -Recurse -File -Force

# Load & parse .picasa.ini files
$picasaIniFiles = $files | Select-Object @{l='IniFile';e= { $_ }},@{l='ini';e={Get-IniContent $_} }
# First pass filter on those files that don't contain contacts (no contacts = no faces)
$picasaIniFiles = @($picasaIniFiles | Where-Object { Test-PicasaIniContainsContacts $_.ini })
# Pull out the sections we need
# Note: Powershell defaults to iterating dictionary values. Since we need the keys as well, we
# have to workaround the default behavior by directly calling GetEnumerator()
$picasaIniFiles = $picasaIniFiles | Select-Object IniFile,@{l='Ini';e={$_.Ini.GetEnumerator()}}, @{l='Contacts';e={$_.Ini[$contactSecion]}}

# Extract all the 'face' INI file entries, carrying along enough information to
# create XMP equivalent metadata

# Flatten into individual section entries. E.g. '[Image21.jpg]'
$iniSections = $picasaIniFiles | Select-Object -Property IniFile,Contacts -ExpandProperty Ini
$iniSections = $iniSections | Select-Object `
    Contacts, `
    @{l='SectionHeader';e={$_.Key}}, `
    @{l='ImageFile';e={Join-Path (Split-Path $_.IniFile  ) $_.Key}},
    @{l='Entries';e={$_.Value.GetEnumerator()}}
# Filter sections that don't refer to a real file
$iniSections = @($iniSections | Where-Object { Test-Path $_.ImageFile })

# Grab ini lines & remove everything but face entries (lines)
$faceEntries = @($iniSections | 
    Select-Object Contacts,ImageFile -ExpandProperty Entries | 
    Where-Object { $_.Key -eq 'faces' })
# Extract & flatten the faces (possible to have multiple faces per ini entry.)
# E.g. faces=rect64(3d12746e6cd8c0e6),ddb135458b46bed4;rect64(9dbe7555ca97bcfe),154a1d1e4ed5fa6a
$faceEntries = $faceEntries | Select-Object Contacts,ImageFile,@{l='Faces';e={$_.Value.Split(';')}}
$faceEntries = $faceEntries | Select-Object Contacts,ImageFile -ExpandProperty Faces
# Parse the individual faces. E.g. rect64(9dbe7555ca97bcfe),154a1d1e4ed5fa6a
$faceEntries = $faceEntries | Select-Object `
    Contacts,ImageFile, `
    @{l='Match';e={Select-String -InputObject $_ -Pattern 'rect64\((.+)\),(.+)'}}
# Filter out failed matches
$faceEntries = @($faceEntries | Where-Object { 
    $_.Match.Matches.Success -and $_.Match.Matches.Groups.Count -eq 3
})
# Extract contact id & filter out invalid contacts
$faceEntries = @($faceEntries | 
    Select-Object *,@{l='ContactId';e={$_.Match.Matches.Groups[2].Value}} |
    Where-Object { $_.Contacts.Contains($_.ContactId) })
# Convert to the structure we need 
$faceEntries = $faceEntries | Select-Object -Property ImageFile, `
    @{l='Rectangle';e={ $_.Match.Matches.Groups[1].Value }}, # The area of the file that contains the face
    @{l='Contact';e={$_.Contacts[$_.ContactId].Replace(';','')}} # The name of the person
# Group the 'face' entries by file, since the metadata is per file. 
$imageFiles = $faceEntries | Group-Object -Property ImageFile

# Convert to Xml
$xmp = $imageFiles | Select-Object -Property Name,@{l='Xml';e={ ConvertTo-Xmp $_.Group }}

# Now apply the XMP metadata - via exiftool.
$exifTool = "C:\Scratch\Image_Processing\exiftool\exiftool.exe"
$scratchXmpFile = "C:\Scratch\test.xmp"

foreach ($entry in $xmp) {
    $entry.Name

    $entry.Xml.Save($scratchXmpFile)
    & $exifTool -overwrite_original -q -q -tagsFromFile $scratchXmpFile $entry.Name
}