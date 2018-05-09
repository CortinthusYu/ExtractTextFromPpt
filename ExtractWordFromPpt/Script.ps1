#--------------------------------------------------------------------------------- 
#The sample scripts are not supported under any Microsoft standard support 
#program or service. The sample scripts are provided AS IS without warranty  
#of any kind. Microsoft further disclaims all implied warranties including,  
#without limitation, any implied warranties of merchantability or of fitness for 
#a particular purpose. The entire risk arising out of the use or performance of  
#the sample scripts and documentation remains with you. In no event shall 
#Microsoft, its authors, or anyone else involved in the creation, production, or 
#delivery of the scripts be liable for any damages whatsoever (including, 
#without limitation, damages for loss of business profits, business interruption, 
#loss of business information, or other pecuniary loss) arising out of the use 
#of or inability to use the sample scripts or documentation, even if Microsoft 
#has been advised of the possibility of such damages 
#--------------------------------------------------------------------------------- 

#requires -version 2.0

Function ConvertTo-OSCWord
{
<#
 	.SYNOPSIS
        ConvertTo-OSCWord is an advanced function which can be used to covert PowerPoint presentation to Word document.
    .DESCRIPTION
        ConvertTo-OSCWord is an advanced function which can be used to covert PowerPoint presentation to Word document.
    .PARAMETER  <Path>
		Specifies the path of slide.
    .EXAMPLE
        C:\PS> ConvertTo-OSCWord -Path D:\PPT\
		File_Name                               Action(Convert to Word)
		---------                               --------------------
		Microsoft1.pptx                         Finished
		Microsoft2.pptx                         Finished
		Microsoft3.pptx                         Finished
		Microsoft4.pptx                         Finished
    .EXAMPLE
        C:\PS> ConvertTo-OSCWord -Path D:\PPT\Microsoft1.pptx
		File Name                               Action(Convert to Word)
		---------                               --------------------
		Microsoft1.pptx                         Finished
#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param
	(
		[Parameter(Mandatory=$true,Position=0)]
		[Alias('p')][String]$Path
	)
	
	if($PSCmdlet.ShouldProcess("Convert PowerPoint presentation to Word document."))
	{
		if(Get-Process|Where-Object {$_.Name -eq "POWERPNT"})
		{
			Get-Process|Where-Object{$_.Name -eq "POWERPNT"}|Stop-Process
			Write-Warning "Some of PowerPoint Applications are running. Please close all PowerPoint Applications that are running and try again."
		}
		else
		{
			if(Test-Path -Path $Path)
			{
				#get all related to powerpoint files
				$PowerPointFiles = Get-ChildItem -Path $Path -Recurse -Include *.ppt,*.pptx,*.pptm,*.ppsx,*.pps,*.ppsm,*.potx,*.pot,*.potm,*.odp
				if($PowerPointFiles)
				{
					#Create the PowerPoint application object
					$PowerPointApp = New-Object -ComObject PowerPoint.Application

					#Create the Word application object
					$WordAPP = New-Object -ComObject Word.Application
					#$WordAPP.Visible=True|Out-Null
					foreach($file in $PowerPointFiles)
					{
						$Objs = @()
						$fileName = $file.Name #get the file name
						$filePath = $file.DirectoryName #get the directory of file
						$fileBaseName = $file.BaseName
						
						$PSCmdlet.WriteVerbose("Opening $file.FullName file")
						#open the slide file in the background.
						$Presentation = $PowerPointApp.Presentations.Open($file.FullName,$null,$null,[Microsoft.Office.Core.MsoTriState]::msoFalse)
						#get all the slides
						$Slides = $Presentation.Slides
						$SlidesCount = $Presentation.Slides.Count
						#get height and width of slides
						#$sldHeight = $Presentation.PageSetup.SlideHeight
						#$sldWidth = $Presentation.PageSetup.SlideWidth
						
						#set default variable
						$wdPaperCustom = 41
						$intPageNumber = 1
						$wdGoToNext = 2
						$wdGoToPage = 1

						#setup word document page size
						$Word=$WordAPP.Documents.Add()
						
						#set the word document size
						#$WordAPP.Selection.PageSetup.LeftMargin = 0
						#$WordAPP.Selection.PageSetup.RightMargin = 0
						#$WordAPP.Selection.PageSetup.TopMargin = 0
						#$WordAPP.Selection.PageSetup.BottomMargin = 0
						$WordAPP.Selection.PageSetup.PaperSize = $wdPaperCustom
						$Selection=$WordAPP.Selection
						#$WordAPP.Selection.PageSetup.PageWidth = $sldWidth
						#$WordAPP.Selection.PageSetup.PageHeight = $sldHeight
						
						foreach($Slide in $Slides)
						{
							Try
							{
								Write-Progress -Activity "Converting PowerPoint presentation [$fileName] to Word" `
								-Status "$intPageNumber of $SlidesCount Slide - Finished" -PercentComplete $($intPageNumber/$SlidesCount*100)

								foreach($Shape in $Slide.Shapes)
								{
									if($Shape.HasTextFrame)
									{
										if($Shape.TextFrame.TextRange.Text -ne "")
										{
											$Text=$Shape.TextFrame.TextRange.Text
											$Selection.TypeText($Text)
											$Selection.TypeParagraph()
										}
									}
								}
								
							#	$SlideRange = $Slide.Shapes.Range()
							#	$SlideRange.Copy()
							#	$WordAPP.Selection.Paste()
							#	$intPageNumber++ #set the page number paramter

							#	$WordAPP.Selection.ShapeRange.Group()|Out-Null
							#	$WordAPP.Selection.ShapeRange.Left = 0
							#	$WordAPP.Selection.ShapeRange.Ungroup()|Out-Null
								
							#	if($NotesText -ne "")
							#	{
							#		$WordAPP.Selection.GoTo($wdGoToPage,$wdGoToNext,$null,$intPageNumber)|Out-Null
							#		#Set the selection text to space,to avoid slide notes overlap.
							#		$WordAPP.Selection.Text = " "
							#		$WordAPP.Selection.MoveRight(1,1)|Out-Null
							#		$Word.Comments.Add($WordAPP.Selection.Range,$NotesText)|Out-Null
							#	}
							#	$WordAPP.Selection.EndKey(6)|Out-Null
							#	$WordAPP.Selection.InsertNewPage()#insert new page to word
							}
							Catch
							{
								Write-Warning "A few presentations have been lost in this converting. NO.$intPageNumber slide cannot convert to word document."
							}
						}
                        
                        $Presentation.Close()
						
						#to delete the last blank page in Word document.
						$WordAPP.Selection.TypeBackspace()
						$WordAPP.Selection.TypeBackspace()
						
						$Word.SaveAs([REF]"$filePath\$fileBaseName")
						$Word.Close()
                
						$Properties = @{'File Name' = $fileName
										'Convert to Word' = if(Test-Path -Path "$filePath\$fileBaseName.docx")
															{"Finished"}
															else
															{"Unfinished"}
										}		
						$objWord = New-Object -TypeName PSObject -Property $Properties
						$Objs += $objWord
						$Objs
					}
					
					#Clear the legacy value of clipboard
					[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
					if ( [System.Windows.Forms.Clipboard]::GetText() -ne ""  )     
					{     
						[System.Windows.Forms.Clipboard]::Clear()
					}    
					
					######release the object######
					$PowerPointApp.Quit() #release the object
					$WordAPP.Quit() #release the object
					[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Presentation)
					[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($PowerPointApp)
					[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word)
					[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp)
					[GC]::Collect()
					[GC]::WaitForPendingFinalizers()
					#close the legacy process
					Get-Process|Where-Object{$_.Name -eq "POWERPNT"}|Stop-Process
				}
				else
				{
					Write-Warning "Please make sure that at least one PowerPoint file in the ""$Path""."
				}
			}
			else
			{
				Write-Warning "The path does not exist, plese input the correct path."
			}
		}
	}
}
ConvertTo-OSCWord -Path .\
