# get-FocalLengths
# Version 7
# Written by Steven Askwith
# stevenaskwith.com
# 18/03/2012

# setup default parameters if none were specified
param([string]$path = (get-location), # current path 
	[string]$fileType = ".jpg", # search for .jpg
	[string]$model = "Canon EOS 550D") # default camera type

##### Functions
function get-FocalLengths ($files)
{
	##### Assemblies 
	# load the .NET Assembly we will be using
	Add-Type -AssemblyName System.Drawing
	
	##### Constants
	$Encode = new-object System.Text.ASCIIEncoding
	# How many files we are working with
	$totalFiles = $files.count
	
	##### Varibles 
	$image = $null
	$imageHash = @{}
	$i = 0
	$focalLength = $null 
	
	foreach ($file in $files)
	{
		# load image by statically calling a method from .NET
		$image = [System.Drawing.Imaging.Metafile]::FromFile($file.FullName)
		
		# try to get the ExIf data (silently fail if the data can't be found)
		# http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
		try
		{
			# get the Focal Length from the Metadata code 37386
			$focalLength = $image.GetPropertyItem(37386).Value[0]
			# get model data from the Metadata code 272
			$modelByte = $image.GetPropertyItem(272)
			# convert the model data to a String from a Byte Array
			$imageModel = $Encode.GetString($modelByte.Value)
			# unload image
			$image.Dispose()	
		}
		catch
		{
			#do nothing with the catch
		}
		
		# if the file contained both focalLength and A modelName
		if(($focalLength -ne $null) -and ($imageModel -eq $model))
		{
			if($imageHash.containsKey($focalLength))
			{
				# incriment count by 1 if focal length is already in hash table
				$count = $imageHash.Get_Item($focalLength)
				$count++
				$imageHash.Set_Item($focalLength,$count)
			}
			else
			{
				# Add focal length to Hash Table if it doesn't exist
				$imageHash.add($focalLength,1) 
			}
		}
		
		# Calculate the current percentage complete
		$i++
		$percentComplete = [math]::round((($i/$totalFiles) * 100), 0)
		
		# Update that lovely percentage bar...
		Write-Progress -Activity:"Loading Focal Lengths" -status "$i of $totalFiles Complete:" -PercentComplete $percentComplete 
	}
	# print results in ascending order of focal length
	return $imageHash
}

function createChart([hashtable]$chartData,[String]$title,[String]$xTitle,[String]$yTitle)
{
	# load the .NET Assembly we will be using
	Add-Type -AssemblyName System.Windows.Forms 
	$assembly = [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	if($assembly -eq $null)
	{
		Write-Host "Charting module not installed, please install it"
		# launch IE and navigate to the correct page
		$ie = New-Object -ComObject InternetExplorer.Application
 		$ie.Navigate("http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=14422")
 		$ie.Visible = $true
		break
	}
	
	# create chart object 
	$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
	$Chart.Width = 500 
	$Chart.Height = 400 
	$Chart.Left = 40 
	$Chart.Top = 30
	 
	# create a chartarea to draw on and add to chart 
	$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
	$Chart.ChartAreas.Add($ChartArea)
	
	# add data to chart 
	[void]$Chart.Series.Add("Data") 
	$Chart.Series["Data"].Points.DataBindXY($chartData.Keys, $chartData.Values)
	
	# add title and axes labels 
	[void]$Chart.Titles.Add($title) 
	$ChartArea.AxisX.Title = $xTitle 
	$ChartArea.AxisY.Title = $yTitle
	
	# Find point with max/min values and change their colour 
	$maxValuePoint = $Chart.Series["Data"].Points.FindMaxByValue() 
	$maxValuePoint.Color = [System.Drawing.Color]::Red 

	$minValuePoint = $Chart.Series["Data"].Points.FindMinByValue() 
	$minValuePoint.Color = [System.Drawing.Color]::Green
	
	# change chart area colour 
	$Chart.BackColor = [System.Drawing.Color]::Transparent
	
	# make bars into 3d cylinders 
	$Chart.Series["Data"]["DrawingStyle"] = "Cylinder"
	
	# display the chart on a form 
	$Chart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left 
	$Form = New-Object Windows.Forms.Form 
	$Form.Text = "PowerShell Chart" 
	$Form.Width = 600 
	$Form.Height = 600 
	$Form.controls.add($Chart) 
	$Form.Add_Shown({$Form.Activate()}) 
	$Form.ShowDialog()
	
	# save chart to file 
	$Chart.SaveImage($path + "\Chart.png", "PNG")
}

##### Main
# Clear the Screen
clear

# clean up the search filter
$filter = "*" + $fileType 
# find all the image files we are interested in
$imageFiles = get-childitem -recurse $path -filter $filter 


# if some files were returned
if ($imageFiles -ne $null)
{
	$imageFreqDist = get-FocalLengths $imageFiles 
	#$imageFreqDist.GetEnumerator() | Sort-Object Name
	createChart $imageFreqDist "Focal Length Frequency Distribution of $($imageFiles.count) Photos" "Focal Length" "Frequency" | Out-Null
}
else
{
	Write-Host "No files found"
}

