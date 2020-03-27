Param
	(
		[String]$Restart
	)



# Add assemblies for WPF and Mahapps
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')    | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')   | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')    | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null
[System.Reflection.Assembly]::LoadFrom("assembly\MahApps.Metro.dll") | out-null
# [System.Reflection.Assembly]::LoadFrom("$Current_Folder\assembly\MahApps.Metro.dll") | out-null


If ($Restart -ne "") 
	{
		sleep 10
	} 
	
$Script:Current_Folder = split-path $MyInvocation.MyCommand.Path
	

 
# Choose an icon to display in the systray
# $icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\wusa.exe") 
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\UserAccountControlSettings.exe") 

########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
#																						INTERFACES A AFFICHER
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################



# ----------------------------------------------------
# Partie Utilisateurs
# ----------------------------------------------------

[xml]$XAML_OK =  
@"
<Controls:MetroWindow 
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"		
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	Title="Config Manager Update Actions" Width="470" ResizeMode="NoResize" Height="300" 
	BorderBrush="DodgerBlue" BorderThickness="0.5" WindowStartupLocation ="CenterScreen">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="$Current_Folder\resources\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

   <Controls:MetroWindow.LeftWindowCommands>
        <Controls:WindowCommands>
            <Button>
                <StackPanel Orientation="Horizontal">
                    <Rectangle Width="15" Height="15" Fill="{Binding RelativeSource={RelativeSource AncestorType=Button}, Path=Foreground}">
                        <Rectangle.OpacityMask>
                            <VisualBrush Stretch="Fill" Visual="{StaticResource appbar_server}" />							
                        </Rectangle.OpacityMask>
                    </Rectangle>					
                </StackPanel>
            </Button>				
        </Controls:WindowCommands>	
    </Controls:MetroWindow.LeftWindowCommands>		

    <Grid>	
	
	<DockPanel>
		<StatusBar DockPanel.Dock="Bottom" Name="statusBar">
			<DockPanel LastChildFill="True" Width="{Binding ActualWidth, ElementName=statusBar, Mode=OneWay}">
				<Label Name="Label_Close_OK" Foreground="White" Content="The window will close in 7 seconds" HorizontalContentAlignment="Center"/>
			</DockPanel>
		</StatusBar>	
		<Menu DockPanel.Dock="Top"/>
	</DockPanel> 	
	
		<StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,-20,0,0">
			<Ellipse Height="80" Width="80">
				<Ellipse.Fill>
					<ImageBrush x:Name="Logo_CM" ImageSource="$current_folder\CM_logo.png"  AlignmentX="Center" AlignmentY="Center" />
				</Ellipse.Fill>
			</Ellipse>			
			<Label FontSize="14" Name="OK_Text" Content="Your Configuration Manager actions have been successfully updated"/>
		</StackPanel>
    </Grid>
</Controls:MetroWindow>        
"@
$OK_Form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML_OK))
$Logo_CM = $OK_Form.findname("Logo_CM") 
$Label_Close_OK = $OK_Form.findname("Label_Close_OK") 
$OK_Text = $OK_Form.findname("OK_Text") 


# ----------------------------------------------------
# Partie Utilisateurs
# ----------------------------------------------------

[xml]$XAML_Warning =  
@"
<Controls:MetroWindow 
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"		
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	Title="Config Manager Update Actions" Width="470" ResizeMode="NoResize" Height="300" 
	BorderBrush="Red" BorderThickness="0.5" WindowStartupLocation ="CenterScreen">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="$Current_Folder\resources\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Red.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

   <Controls:MetroWindow.LeftWindowCommands>
        <Controls:WindowCommands>
            <Button>
                <StackPanel Orientation="Horizontal">
                    <Rectangle Width="15" Height="15" Fill="{Binding RelativeSource={RelativeSource AncestorType=Button}, Path=Foreground}">
                        <Rectangle.OpacityMask>
                            <VisualBrush Stretch="Fill" Visual="{StaticResource appbar_warning}" />							
                        </Rectangle.OpacityMask>
                    </Rectangle>					
                </StackPanel>
            </Button>				
        </Controls:WindowCommands>	
    </Controls:MetroWindow.LeftWindowCommands>		

    <Grid>	
	
	<DockPanel>
		<StatusBar DockPanel.Dock="Bottom" Name="statusBar">
			<DockPanel LastChildFill="True" Width="{Binding ActualWidth, ElementName=statusBar, Mode=OneWay}">
				<Label Name="Label_Close_KO" Foreground="White" Content="The window will close in 7 seconds" HorizontalContentAlignment="Center"/>
			</DockPanel>
		</StatusBar>	
		<Menu DockPanel.Dock="Top"/>
	</DockPanel> 

		<StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,-10,0,0">
			<Ellipse Height="80" Width="80">
				<Ellipse.Fill>
					<ImageBrush x:Name="Logo_CM" ImageSource="$current_folder\CM_logo_KO.png" AlignmentX="Center" AlignmentY="Center" />
				</Ellipse.Fill>
			</Ellipse>			
			<Label Margin="0,10,0,0" HorizontalAlignment="Center" FontSize="14" Foreground="Red" FontWeight="Bold" Content="Oops something went wrong"/>

			<StackPanel Name="Warning_Before_Update">
				<TextBlock FontSize="14" Name="Message_Error_Warning" TextWrapping="Wrap"  TextAlignment="Center"/>		
			</StackPanel>				
			
			<StackPanel Name="Warning_During_Update">
			<!--	<ScrollViewer  CanContentScroll="True" Height="80">    -->    										
				<!--	<TextBlock FontSize="14" x:Name="Message_Error_Warning_Update" TextWrapping="Wrap" /> -->
					<Button	x:Name="Message_Error_Warning_Update" Content="View details" Width="80" Height="20"/>
			<!--	</ScrollViewer> -->
			</StackPanel>			
		</StackPanel>
    </Grid>
</Controls:MetroWindow>     
"@
$Warning_Form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML_Warning))
$Label_Close_KO = $Warning_Form.findname("Label_Close_KO") 
$Message_Error_Warning = $Warning_Form.findname("Message_Error_Warning") 
$Warning_Before_Update = $Warning_Form.findname("Warning_Before_Update") 
$Warning_During_Update = $Warning_Form.findname("Warning_During_Update") 
$Message_Error_Warning_Update = $Warning_Form.findname("Message_Error_Warning_Update") 


 
# Add icon the systray 
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "Configuration Manager Force Update"
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true


# Add Restart the tool
$Menu_Force_Update = New-Object System.Windows.Forms.MenuItem
$Menu_Force_Update.Text = "CM Force Update"

# Add Restart the tool
$Menu_Restart_Tool = New-Object System.Windows.Forms.MenuItem
$Menu_Restart_Tool.Text = "Restart the tool (in 10secs)"
 
# Add menu exit
$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"
 
# Add all menus as context menus
$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Force_Update)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Restart_Tool)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)

$Tool_log_file = "$Temp_Folder\CMCB_Refresh_Tool.log"
If(!(test-path $Tool_log_file))
	{
		new-item $Tool_log_file -force -type file		
	}

function Get-Time {   
    return "{0:MM/dd/yy} {0:HH:mm:ss}" -f (Get-Date)   
}	

Function Pre_Check_Part
	{	
		################# Check services and Site code #################
		################# Test presence of SCCM	#################
		add-content $Tool_log_file  "-----------------------------------------------------------------------------------------------------------------------------------"				
		add-content $Tool_log_file  "$(Get-Time) - Prerequisites part 2 - Check CMCB Client"	
		add-content $Tool_log_file  "-----------------------------------------------------------------------------------------------------------------------------------"		

		# $SCCM_Test = Test-Path 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client'
		$Client_status.Fontweight = "bold"	
		$Site_code.Fontweight = "bold"		

		$SMS_Agent_Service_Name = 'CcmExec'
		$SMS_Agent_Service = Get-Service -Name $SMS_Agent_Service_Name
		If ($SMS_Agent_Service -ne $null)
			{
				If ($SMS_Agent_Service.status -ne "Running")
					{
						add-content $Tool_log_file  "$(Get-Time) - SMS Agent Host service is not running"						
						$Script:Client = "KO"
						$Script:Warning_Message = "Service not started"
					}
				Else
					{
						add-content $Tool_log_file  "$(Get-Time) - SMS Agent Host service is started"											
						$Script:Client = "OK"
						$Script:Warning_Message = "Service started"		
					}
								
				$REG_SCCM = get-itemproperty -path registry::"HKLM\SOFTWARE\Microsoft\SMS\Mobile Client"	
				$SCCM_Site_Code = $REG_SCCM.AssignedSiteCode		
				$Site_code.Content = $SCCM_Site_Code
				$Script:Site = "OK"						
				add-content $Tool_log_file  "$(Get-Time) - Site code is $SCCM_Site_Code"										
				If ($SCCM_Site_Code -eq $null)
					{						
						add-content $Tool_log_file  "$(Get-Time) - No site code is assigned"											
						$Script:Site = "KO"
						$Script:Warning_Message = "No site assigned"				
					}				
			}
		Else
			{
				add-content $Tool_log_file  "$(Get-Time) - SMS Agent Host service can not be found. Client seems not to be installed"													
				$Script:Client = "KO"
				$Script:Warning_Message = "Client not installed"
			}	
	

		################ Check network connection ###############
		add-content $Tool_log_file  "$(Get-Time) - Prerequisites part 1 - Check network connection"	
		add-content $Tool_log_file  "-----------------------------------------------------------------------------------------------------------------------------------"				
						
		################ Check network connection: If ping to the MP is OK ###############
		$REG_CCMSetup = "HKLM:SOFTWARE\Microsoft\CCM"
		$Test_REG_CCMSetup = test-path $REG_CCMSetup
		If ($Test_REG_CCMSetup)
			{
				$CCMSetup_Get_LastValidMP = (gwmi -namespace root\ccm -class sms_authority).CurrentManagementPoint

				If ($CCMSetup_Get_LastValidMP -ne $null)
					{						
						add-content $Tool_log_file  "$(Get-Time) - Management Point is $Get_LastValidMP_Value"	
					
						$Test_Connection_LastValidMP = Test-connection $CCMSetup_Get_LastValidMP -count 3 -quiet
						If ($Test_Connection_LastValidMP)
							{
								add-content $Tool_log_file  "$(Get-Time) - Connection to the Management Point is OK"								
								$Script:Network = "OK"								
							}
						Else
							{
								add-content $Tool_log_file  "$(Get-Time) - Connection to the Management Point is KO"															
								$Script:Network = "KO"							
							}						
					}
				Else
					{
						add-content $Tool_log_file  "$(Get-Time) - No Management Point found"													
						$Script:Network = "KO"														
					}			
			}
		Else
			{
				add-content $Tool_log_file  "$(Get-Time) - The following registry key can't be found: HKLM:SOFTWARE\Microsoft\CCM"																
				$Script:Network = "KO"		
			}


		# If site is OK and connection to the MP is ok we change the first line color			
		If (($Client -eq "OK") -and ($Site -eq "OK") -and ($Network -eq "OK"))
			{
				$Script:Warning_Message = "Updates requested"	
			}
		Else
			{  													
				If (($Client -eq "KO") -and ($Site -eq "KO") -and ($Network -eq "KO")) 
					{
						$Message_Error_Warning.Text = "The SMS Agent Host service isn't running or can't be found`n"
						$Message_Error_Warning.Text += "There is no site assigned`n"
						$Message_Error_Warning.Text += "Can not connect to the Management Point"			
					}
				
				If (($Client -eq "KO") -and ($Site -eq "KO") -and ($Network -eq "OK")) 
					{
						$Message_Error_Warning.Text = "The SMS Agent Host service isn't running or can't be found`n"
						$Message_Error_Warning.Text += "There is no site assigned"			
					}				
				
				
				If (($Client -eq "KO") -and ($Site -eq "OK") -and ($Network -eq "KO")) 
					{
						$Message_Error_Warning.Text = "The SMS Agent Host service isn't running or can't be found`n"
						$Message_Error_Warning.Text += "Can not connect to the Management Point"										
					}					
				
				
				If (($Client -eq "KO") -and ($Site -eq "OK") -and ($Network -eq "OK")) 
					{
						$Message_Error_Warning.Text = "The SMS Agent Host service isn't running or can't be found"															
					}	

					
				If (($Client -eq "OK") -and ($Site -eq "KO") -and ($Network -eq "KO")) 
					{
						$Message_Error_Warning.Text = "There is no site assigned`n"
						$Message_Error_Warning.Text += "Can not connect to the Management Point"			
					}
				
				If (($Client -eq "OK") -and ($Site -eq "KO") -and ($Network -eq "OK")) 
					{						
						$Message_Error_Warning.Text = "There is no site assigned"				
					}				
				
				
				If (($Client -eq "OK") -and ($Site -eq "OK") -and ($Network -eq "KO")) 
					{
						$Message_Error_Warning.Text = "Can not connect to the Management Point"										
					}					
			}

	# $Script:Client = "OK"
	# $Script:Site = "OK"
	# $Script:Network = "OK"
			
}	



Function Run_Actions
	{
		add-content $Tool_log_file  "-----------------------------------------------------------------------------------------------------------------------------------"						
		add-content $Tool_log_file  "$(Get-Time) - Run actions Part"	
		add-content $Tool_log_file  "-----------------------------------------------------------------------------------------------------------------------------------"				

		# $Config_Manager_Actions_xml = "$CM_Config_Folder\Config_Manager_Actions.xml"
		$Config_Manager_Actions_xml = "$Current_Folder\Config_Manager_Actions.xml"
		$Config_Manager_Object = New-Object -ComObject CPApplet.CPAppletMgr			
		$Global:list_Actions = $Config_Manager_Actions_xml						
		$Get_Actions = [xml] (Get-Content $list_Actions)			
		$Actions = $Get_Actions.Actions.Action | Where {$_.Active -match "True"}


		$Script:message_action = @()
		$Script:Issue_During_Update = $False
		foreach ($action in $Actions) 
			{	
			
				$action_name = $action.Name	
				
				$Config_Manager_Actions = $Config_Manager_Object.GetClientActions() | Where-Object {($_.Name -eq $action_name)}			
				Foreach ($action_To_Run in $Config_Manager_Actions)
					{		
						$action_To_Run_Name = $action_To_Run.Name
					
						try 
							{
								# $action_To_Run.PerformAction()	
								$Config_Manager_Actions.PerformAction()										
								$Action_Status = "Successfully executed $action_name"	
								write-host $Action_Status
								$Script:message_action += "$Action_Status `n"
								add-content $Tool_log_file  "$(Get-Time) - $Action_Status"	
							} 
						catch 
							{
								$Action_Status = "Failed to execute $action_name"	
								write-host $Action_Status
								
								$Script:message_action += "$Action_Status `n"
								$Script:Issue_During_Update = $True
								add-content $Tool_log_file  "$(Get-Time) - $Action_Status"	
							}	
					
						break								
					}
			}				
	}




Function Success_Form
	{
		$Script:Timer = New-Object System.Windows.Forms.Timer
		$Timer.Interval = 1000

		Function Timer_Tick()
		{
			$Label_Close_OK.Content = "The window will close in $Script:CountDown seconds"
			 --$Script:CountDown
			 If ($Script:CountDown -lt 0)
			 {
				ClearAndClose
				$Timer.Stop(); 
				$OK_Form.Close(); 
				$Timer.Dispose();
				$Script:CountDown = 5		
			 }
		}

		$Script:CountDown = 6
		$Timer.Add_Tick({ Timer_Tick})
		$Timer.Start()	
		$OK_Form.ShowDialog()
		$OK_Form.Activate() 	
	}
	
Function Warning_Form
	{
		$Script:Timer = New-Object System.Windows.Forms.Timer
		$Timer.Interval = 1000

		Function Timer_Tick()
		{
			$Label_Close_KO.Content = "The window will close in $Script:CountDown seconds"
			 --$Script:CountDown
			 If ($Script:CountDown -lt 0)
			 {
				ClearAndClose
				$Timer.Stop(); 
				$Warning_Form.Close(); 
				$Timer.Dispose();
				$Script:CountDown = 5		
			 }
		}

		$Script:CountDown = 6
		$Timer.Add_Tick({ Timer_Tick})
		$Timer.Start()	
		$Warning_Form.ShowDialog()
		$Warning_Form.Activate() 	
	}




# Action after a click on the main systray icon
$Main_Tool_Icon.Add_Click({     
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Interface_Utilisateurs)
	If ($_.Button -eq [Windows.Forms.MouseButtons]::Left) 
		{
			$Interface_Utilisateurs.WindowStartupLocation = "CenterScreen"
			$Interface_Utilisateurs.Show()
			$Interface_Utilisateurs.Activate() 

			Pre_Check_Part
			$Warning_During_Update.Visibility = "Collapsed"	

			If(($Client -eq "OK") -and ($Site -eq "OK") -and ($Network -eq "OK"))
				{
				
					Run_Actions
					If($Issue_During_Update -eq $True)
						{
							$Warning_During_Update.Visibility = "Visible"
							$Message_Error_Warning_Update_Text.Text = $message_action	
							Warning_Form										
						}
					Else	
						{
							$Warning_During_Update.Visibility = "Collapsed"	
							$Message_Error_Warning_Update_Text.Text = ""						
							Success_Form				
						}
			
				}
			Else
				{
					Warning_Form	
				}			

		}    
})
 

$Warning_Form.Add_Closing({
	$Timer.Stop(); 
	$Timer.Dispose();	
	$_.Cancel = $true
	$Warning_Form.Hide()	
})

$OK_Form.Add_Closing({
	$Timer.Stop(); 
	$Timer.Dispose();	
	$_.Cancel = $true
	$OK_Form.Hide()	
})





$Menu_Force_Update.add_Click({

	Pre_Check_Part
	$Warning_During_Update.Visibility = "Collapsed"	

	If(($Client -eq "OK") -and ($Site -eq "OK") -and ($Network -eq "OK"))
		{		
			Run_Actions
			If($Issue_During_Update -eq $True)
				{
					$Warning_During_Update.Visibility = "Visible"
					$Message_Error_Warning_Update_Text.Text = $message_action	
					Warning_Form										
				}
			Else	
				{
					$Warning_During_Update.Visibility = "Collapsed"	
					$Message_Error_Warning_Update_Text.Text = ""						
					Success_Form				
				}
	
		}
	Else
		{
			Warning_Form	
		}
})

# ===============================================
# ======== Load TemplateWindow ==================
# ===============================================

function LoadXml ($global:filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}
$xamlDialog  = LoadXml("Dialog.xaml")
$read=(New-Object System.Xml.XmlNodeReader $xamlDialog)
$DialogForm=[Windows.Markup.XamlReader]::Load( $read )
$CustomDialog = [MahApps.Metro.Controls.Dialogs.CustomDialog]::new($Warning_Form)
$CustomDialog.AddChild($DialogForm)

$Cancel = $DialogForm.FindName("Cancel")
$Message_Error_Warning_Update_Text = $DialogForm.FindName("Message_Error_Warning_Update_Text")

$Cancel.add_Click({
    $CustomDialog.RequestCloseAsync()
})









$Message_Error_Warning_Update.add_Click({
	$Message_Error_Warning_Update_Text.Text = $message_action		
	[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Warning_Form, $CustomDialog)
})

# Action after clicking on the Restart context menu
$Menu_Restart_Tool.add_Click({
	$Restart = "Yes"
	start-process -WindowStyle hidden powershell.exe "$Current_Folder\CM_Force_Update.ps1 '$Restart'" 	

	$Main_Tool_Icon.Visible = $false
	$window.Close()
	Stop-Process $pid	
 })
 
 
# Action after clicking on the Exit context menu
$Menu_Exit.add_Click({
	$Main_Tool_Icon.Visible = $false
	$window.Close()
	Stop-Process $pid
})

# Make PowerShell Disappear - Thanks Chrissy
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
 
# Use a Garbage colection to reduce Memory RAM
# https://dmitrysotnikov.wordpress.com/2012/02/24/freeing-up-memory-in-powershell-using-garbage-collector/
# https://docs.microsoft.com/fr-fr/dotnet/api/system.gc.collect?view=netframework-4.7.2
[System.GC]::Collect()
 
# Create an application context for it to all run within - Thanks Chrissy
# This helps with responsiveness, especially when clicking Exit - Thanks Chrissy
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)


