Add-Type -assembly System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Core | Out-Null

<####### Назначение основной формы #######>

$Window = New-Object System.Windows.Forms.Form
$Window.MinimumSize = New-Object System.Drawing.Size (885, 757);
$Window.MaximumSize = New-Object System.Drawing.Size (885, 757);
#$Window.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13);		
#$Window.ClientSize = New-Object System.Drawing.Size (1008, 729);
$window.StartPosition = "CenterScreen";
$Window.SizeGripStyle = "Hide";
$Window.Text = "HelpDesk GUI Tool";
$Window.AutoScaleMode = "Font";

    $MainBox = New-Object System.Windows.Forms.TabControl;
    $MainBox.Location = New-Object System.Drawing.Point(12, 12);
    $MainBox.Size = New-Object System.Drawing.Size(845, 694);
    $MainBox.Anchor = @("Top", "Right", "Left", "Bottom");
 ###$MainBox.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;

        $UserBox = New-Object System.Windows.Forms.TabPage;
	    $UserBox.Location = New-Object System.Drawing.Point(4, 22);
	    $UserBox.Size = New-Object System.Drawing.Size(976, 679);
	    $UserBox.Padding = New-Object System.Windows.Forms.Padding(3);
        $UserBox.BackColor = "white";
        $UserBox.Text = "Карточка пользователя";		                  		

                $label1 = New-Object System.Windows.Forms.Label;
                $label1.Location = New-Object System.Drawing.Point(15, 15);
			    $label1.Size = New-Object System.Drawing.Size(110, 20);
			    $label1.TextAlign = "MiddleCenter";
			    $label1.Text = "Имя пользователя:";
                $label1.name = 'label1'
            
                $UserSearch_Box = New-Object System.Windows.Forms.ComboBox;
       			$UserSearch_Box.Location = New-Object System.Drawing.Point(131, 16);
			    $UserSearch_Box.Size = New-Object System.Drawing.Size(250, 20);			
                $UserSearch_Box.MaxLength = 24;
                $UserSearch_Box.AllowDrop = $true; 
                $UserSearch_Box.Name = 'UserSearch_Box'
                            
                $UserSearch_But = New-Object System.Windows.Forms.Button;
        		$UserSearch_But.Location = New-Object System.Drawing.Point(387, 15);
			    $UserSearch_But.Size = New-Object System.Drawing.Size(23, 23);
        		$UserSearch_But.Anchor = @("Left", "Top");
			    $UserSearch_But.AutoEllipsis = $true;
			    $UserSearch_But.Cursor = "Hand";
			    $UserSearch_But.Text = "o";
                $UserSearch_But.Name = 'UserSearch_But'
			    $UserSearch_But.UseVisualStyleBackColor = $true;


            $UserInfo = New-Object System.Windows.Forms.GroupBox;
    		$UserInfo.Location = New-Object System.Drawing.Point(15, 43);
			$UserInfo.Size = New-Object System.Drawing.Size(400, 260);
			$UserInfo.BackColor = "WhiteSmoke";
			$UserInfo.TabStop = $false;
			$UserInfo.Text = "Общее";
            
                $U_stat = New-Object System.Windows.Forms.Label;
			    $U_stat.Location = New-Object System.Drawing.Point(150, 5);
			    $U_stat.Size = New-Object System.Drawing.Size(250, 20);
				$U_stat.Anchor = @("Top","Right");
			    $U_stat.BackColor = "Transparent";
			    $U_stat.Visible = $false;
			    $U_stat.TextAlign = "MiddleRight";
			    $U_stat.Text = "(status)";


                $U_fullname = New-Object System.Windows.Forms.textbox;######################################
                $U_fullname.Location = New-Object System.Drawing.Point(8, 25)           
			    $U_fullname.Size = New-Object System.Drawing.Size(360, 20);
                $U_fullname.BackColor = "Whitesmoke"
                $U_fullname.BorderStyle = "none"
                $U_fullname.ReadOnly = $true
			    #$U_fullname.TextAlign = "MiddleLeft";
			    $U_fullname.Text = "(fullname)";
            
                $U_login = New-Object System.Windows.Forms.Label;
            	$U_login.Location = New-Object System.Drawing.Point(6, 65);
			    $U_login.Size = New-Object System.Drawing.Size(360, 20);
			    $U_login.TextAlign = "MiddleLeft";
			    $U_login.Text = "(login)";
            
                $U_id = New-Object System.Windows.Forms.Label;
        		$U_id.Location = New-Object System.Drawing.Point(6, 45);
			    $U_id.Size = New-Object System.Drawing.Size(360, 20);
			    $U_id.TextAlign = "MiddleLeft";
			    $U_id.Text = "(id)";          
            
                $U_email = New-Object System.Windows.Forms.Label;
        		$U_email.Location = New-Object System.Drawing.Point(6, 85);
			    $U_email.Size = New-Object System.Drawing.Size(360, 20);
			    $U_email.TextAlign = "MiddleLeft";
			    $U_email.Text = "(email)";
          
                $U_phones = New-Object System.Windows.Forms.Label;
           		$U_phones.Location = New-Object System.Drawing.Point(6, 105);
			    $U_phones.Size = New-Object System.Drawing.Size(360, 20);
			    $U_phones.TextAlign = "MiddleLeft";
			    $U_phones.Text = "(phones)";

                $label2 = New-Object System.Windows.Forms.Label;
      			$label2.Location = New-Object System.Drawing.Point(6, 128);
			    $label2.Size = New-Object System.Drawing.Size(85, 20);
			    $label2.TextAlign = "MiddleLeft";
			    $label2.Text = "Руководитель:";
                       
                $LinkToBoss = New-Object System.Windows.Forms.LinkLabel;
               	$LinkToBoss.Location = New-Object System.Drawing.Point(86, 128);
			    $LinkToBoss.Size = New-Object System.Drawing.Size(280, 20);
			    $LinkToBoss.TextAlign = "MiddleLeft";
			    $LinkToBoss.TabStop = $true;
			    $LinkToBoss.Text = "(boss)";
			    $LinkToBoss.VisitedLinkColor = "Blue";
                $LinkToBoss.add_Click({ $UserSearch_Box.Text = $LinkToBoss.Text; Search_User })
            
                $label3 = New-Object System.Windows.Forms.Label;           
			    $label3.Location = New-Object System.Drawing.Point(6, 169);
			    $label3.Size = New-Object System.Drawing.Size(88, 20);
			    $label3.TextAlign = "MiddleLeft";
			    $label3.Text = "Создан:";
            
                $U_made = New-Object System.Windows.Forms.Label;
       			$U_made.Location = New-Object System.Drawing.Point(100, 169);
			    $U_made.Size = New-Object System.Drawing.Size(266, 20);
			    $U_made.TextAlign = "MiddleLeft";
			    $U_made.Text = "(u_made)";
            
                $label4 = New-Object System.Windows.Forms.Label;
            	$label4.Location = New-Object System.Drawing.Point(6, 189);
			    $label4.Size = New-Object System.Drawing.Size(88, 20);
			    $label4.TextAlign = "MiddleLeft";
			    $label4.Text = "Изменён:";
            
                $U_changed = New-Object System.Windows.Forms.Label;
            	$U_changed.Location = New-Object System.Drawing.Point(100, 189);
			    $U_changed.Size = New-Object System.Drawing.Size(266, 20);
			    $U_changed.TextAlign = "MiddleLeft";
			    $U_changed.Text = "(u_change)";
                      
                $Uid_name = New-Object System.Windows.Forms.Label;
                $Uid_name.Location = New-Object System.Drawing.Point(6, 209);
			    $Uid_name.Size = New-Object System.Drawing.Size(360, 40);
			    $Uid_name.TextAlign = "MiddleLeft";
			    $Uid_name.Text = "(uid_name)";

            $PassInfo = New-Object System.Windows.Forms.GroupBox;
			$PassInfo.Location = New-Object System.Drawing.Point(15, 309);
			$PassInfo.Size = New-Object System.Drawing.Size(400, 80);
			$PassInfo.TabStop = $false;
			$PassInfo.BackColor = "WhiteSmoke";
			$PassInfo.Text = "Пароль";

                $Pass_stat = New-Object System.Windows.Forms.Label;
			    $Pass_stat.Location = New-Object System.Drawing.Point(150, 5);
			    $Pass_stat.Size = New-Object System.Drawing.Size(250, 20);
                $Pass_stat.Anchor = @("Top","Right");
                $Pass_stat.BackColor = "Transparent";
			    $Pass_stat.Visible = $false;
			    $Pass_stat.TextAlign = "MiddleRight";
			    $Pass_stat.Text = "(pass_stat)";
            
                $label5 = New-Object System.Windows.Forms.Label;
    			$label5.Location = New-Object System.Drawing.Point(6, 25);
			    $label5.Size = New-Object System.Drawing.Size(80, 20);
			    $label5.TextAlign = "MiddleLeft";            
			    $label5.Text = "Установлен:";
            
                $pass_made = New-Object System.Windows.Forms.Label;
				$pass_made.Location = New-Object System.Drawing.Point(86, 25);
			    $pass_made.Size = New-Object System.Drawing.Size(158, 20);
			    $pass_made.TextAlign = "MiddleLeft";
			    $pass_made.Text = "(pass_made)";

                $pass_life = New-Object System.Windows.Forms.Label;
				$pass_life.Location = New-Object System.Drawing.Point(250, 25);
			    $pass_life.Size = New-Object System.Drawing.Size(102, 20);
			    $pass_life.TextAlign = "MiddleLeft";
			    $pass_life.Text = "(pass_life)";

                $PassReset_But = New-Object System.Windows.Forms.Button;
            	$PassReset_But.Location = New-Object System.Drawing.Point(68, 50);
			    $PassReset_But.Size = New-Object System.Drawing.Size(80, 24);
			    $PassReset_But.Text = "Сбросить";
			    $PassReset_But.UseVisualStyleBackColor = $true;
            
                $PassChange_But = New-Object System.Windows.Forms.Button;
            	$PassChange_But.Location = New-Object System.Drawing.Point(154, 50);
			    $PassChange_But.Size = New-Object System.Drawing.Size(80, 24);
			    $PassChange_But.Text = "Изменить";
			    $PassChange_But.UseVisualStyleBackColor = $true;
			
                $LockOutStat_But = New-Object System.Windows.Forms.Button;
            	$LockOutStat_But.Location = New-Object System.Drawing.Point(240, 50);
			    $LockOutStat_But.Size = New-Object System.Drawing.Size(80, 24);
			    $LockOutStat_But.Text = "LockStat";
			    $LockOutStat_But.UseVisualStyleBackColor = $true;
                $LockOutStat_But.Enabled = $false

            $TermsInfo = New-Object System.Windows.Forms.GroupBox
            $TermsInfo.Location = New-Object System.Drawing.Point(15, 395)
            $TermsInfo.Size = New-Object System.Drawing.Point(400, 260)
            $TermsInfo.Text = "Фермы"
            $TermsInfo.BackColor = "WhiteSmoke"
            $TermsInfo.TabStop = $false

                $FindTerms_But = New-Object System.Windows.Forms.Button
                $FindTerms_But.Location = New-Object System.Drawing.Point(6, 19);
	            $FindTerms_But.Size = New-Object System.Drawing.Size(80, 24);
	            $FindTerms_But.Text = "Поиск";
                $FindTerms_But.UseVisualStyleBackColor = $true;
                $FindTerms_But.Enabled = $false

                $UseTH = New-Object System.Windows.Forms.CheckBox
                $UseTH.Location = New-Object System.Drawing.Point(8, 46);
	            $UseTH.Size = New-Object System.Drawing.Size(50, 15);
	            $UseTH.Text = "TH";

                $UseVTS = New-Object System.Windows.Forms.CheckBox
                $UseVTS.Location = New-Object System.Drawing.Point(8, 67);
	            $UseVTS.Size = New-Object System.Drawing.Size(50, 15);
	            $UseVTS.Text = "VTS";

                $TermsLogoff_But = New-Object System.Windows.Forms.Button
                $TermsLogoff_But.Location = New-Object System.Drawing.Point(6, 114);
	            $TermsLogoff_But.Size = New-Object System.Drawing.Size(80, 24);
	            $TermsLogoff_But.Text = "Сброс";
                $TermsLogoff_But.Enabled = $false
                $TermsLogoff_But.UseVisualStyleBackColor = $true;

                $TermsProfile_But = New-Object System.Windows.Forms.Button
                $TermsProfile_But.Location = New-Object System.Drawing.Point(6, 144);
	            $TermsProfile_But.Size = New-Object System.Drawing.Size(80, 24);
	            $TermsProfile_But.Text = "Профиль..";
                $TermsProfile_But.UseVisualStyleBackColor = $true;
                $TermsProfile_But.Enabled = $false

                $TermsRA_But = New-Object System.Windows.Forms.Button
                $TermsRA_But.Location = New-Object System.Drawing.Point(6, 174);
	            $TermsRA_But.Size = New-Object System.Drawing.Size(80, 24);
	            $TermsRA_But.UseVisualStyleBackColor = $true;
                $TermsRA_But.Text = "Помощь"
                $TermsRA_But.Enabled = $false
                
                $PBarTerms = New-Object System.Windows.Forms.ProgressBar
                $PBarTerms.Location = New-Object System.Drawing.Point(92, 19);
	            $PBarTerms.Size = New-Object System.Drawing.Size(95, 24);
                $PBarTerms.Hide()

                $PBarStat = New-Object System.Windows.Forms.Label
                $PBarStat.Location = New-Object System.Drawing.Point(92, 46);
	            $PBarStat.Size = New-Object System.Drawing.Size(95, 20);
	            $PBarStat.TabIndex = 9;
	            $PBarStat.Text = "(PBarStat)";
                $PBarStat.Hide()

                $TermsList = New-Object System.Windows.Forms.TextBox
	            $TermsList.Location = New-Object System.Drawing.Point(193, 19);
                $TermsList.Size = New-Object System.Drawing.Size(200, 235);
                #$TermsList.AcceptsReturn = $true;
	            $TermsList.Multiline = $true;
                $TermsList.ReadOnly = $true
                $TermsList.Hide()

            $GroupsInfo = New-Object System.Windows.Forms.GroupBox
            $GroupsInfo.Location = New-Object System.Drawing.Point(421, 43)
            $GroupsInfo.Size = New-Object System.Drawing.Point(400, 346)
            $GroupsInfo.BackColor = 'WhiteSmoke'
            $GroupsInfo.TabStop = $false;
            $GroupsInfo.Text = 'Группы'

                $GroupsListBox = New-Object System.Windows.Forms.ListBox
                $GroupsListBox.Location = New-Object System.Drawing.Point(6, 19)
                $GroupsListBox.Size = New-Object System.Drawing.Point(390, 316)    
                #$GroupsListBox.ReadOnly = $true

<####### Назначение родительско-дочерних связей. После этих манипуляций форма выходит в свет. #######>

            $Groupsinfo.Controls.Add($GroupsListBox)

            $TermsInfo.Controls.Add($TermsList)
            $TermsInfo.Controls.Add($PBarStat)
            $TermsInfo.Controls.Add($PBarTerms)
            $TermsInfo.Controls.Add($TermsRA_But)
            $TermsInfo.Controls.Add($TermsProfile_But)
            $TermsInfo.Controls.Add($TermsLogoff_But)
            $TermsInfo.Controls.Add($UseVTS)
            $TermsInfo.Controls.Add($UseTH)
            $TermsInfo.Controls.Add($FindTerms_But)

			$PassInfo.Controls.Add($LockOutStat_But);
			$PassInfo.Controls.Add($PassChange_But);
			$PassInfo.Controls.Add($PassReset_But);
			$PassInfo.Controls.Add($pass_life);
			$PassInfo.Controls.Add($pass_made);
			$PassInfo.Controls.Add($label5);
			$PassInfo.Controls.Add($Pass_stat);

			$UserInfo.Controls.Add($Uid_name);
			$UserInfo.Controls.Add($U_changed);
			$UserInfo.Controls.Add($label4);
			$UserInfo.Controls.Add($U_made);
			$UserInfo.Controls.Add($label3);
			$UserInfo.Controls.Add($LinkToBoss);
			$UserInfo.Controls.Add($label2);
			$UserInfo.Controls.Add($U_phones);
			$UserInfo.Controls.Add($U_email);
			$UserInfo.Controls.Add($U_login);
			$UserInfo.Controls.Add($U_id);
			$UserInfo.Controls.Add($U_fullname);
			$UserInfo.Controls.Add($U_stat);
        $UserBox.Controls.Add($GroupsInfo)
        $UserBox.Controls.Add($TermsInfo)
        $UserBox.Controls.Add($PassInfo);
        $UserBox.Controls.Add($UserInfo);
        $UserBox.Controls.Add($UserSearch_But);
        $UserBox.Controls.Add($UserSearch_Box);
        $UserBox.Controls.Add($label1);
    $MainBox.Controls.Add($UserBox);
$Window.Controls.Add($MainBox);

<####### Обработка элементов формы непосредственно перед выводом #######>

$UserBox.controls | %{  if(($_.name -ne 'label1') -and ($_.name -ne 'UserSearch_Box') -and ($_.name -ne 'UserSearch_But')){$_.hide()}}

<####### Подключение библиотеки #######>


Add-PSSnapin Quest.ActiveRoles.ADManagement 2>$null
if (!(Get-PSSnapin Quest.ActiveRoles.ADManagement)){ Show_Alert 0 } 



function Search_User(){
    if ($UserSearch_Box.Text -ne ''){        
        
        #Заполняем историю поиска
        $history = $($UserSearch_Box.Items)
        $UserSearch_Box.Items.Clear()        
        $UserSearch_Box.Items.Add($UserSearch_Box.Text)
        if ($history){ $UserSearch_Box.Items.AddRange($history) }
        
        Connect-QADService -service 'hq-dc.okmarket.ru' | Out-Null        
        
        $u = Get-QADUser $UserSearch_Box.Text -Properties SamAccountName, DisplayName, Email, employeeNumber, telephoneNumber, `
        manager, whenCreated, whenChanged, CanonicalName, PasswordLastSet, PasswordAge, PasswordIsExpired, UserMustChangePassword, AccountIsLockedOut, AccountIsDisabled, memberOf
        $Global:user = $u.SamAccountName
        
        if (!$user) { Show_Alert 1 } 
            elseif (($user.gettype()).name -eq 'Object[]') { Show_Alert 2 }
                elseif (($user.gettype()).name -eq 'String') {
   	               <# $SCCMout = $null
  	                $SQLCon = New-Object System.Data.SQLClient.SQLConnection
	                $SQLSrv = 'central-sccm'
	                $SQLCon.ConnectionString = "Server=$SQLSrv; Integrated Security = True"
	                $SQLCon.Open()
	                $Query='SELECT * FROM [CM_MCM].[dbo].[v_SUP_ConsolUser] where "console use" ='+"'okey\$username'"
	                $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($Query, $SQLCon)
	                $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($SQLCommand)
	                $DataSet = New-Object System.Data.DataSet
	                $SQLAdapter.Fill($DataSet) >> $null
	                $SCCMout = $DataSet.Tables[0] | select "Name PC","Last Console Use" | Sort-Object -Property "Last Console Use" -Descending | where{$_."Name PC" -notlike "vts-*"}
	                $SQLCon.Close()#>
                    
###Вывод инфо       
                    $UserBox.Controls | %{if(!$_.visible){$_.visible = $true}}
                    $U_fullname.Text = $($u.DisplayName)
                    $U_login.Text = $user
                    $U_email.Text = $($u.Email)
                
                    
                    $U_id.Text = $($u.employeeNumber)
                
                    $U_phones.Text = $($u.telephoneNumber)
                
                    $LinkToBoss.text = if($u.manager){$label2.show(); $(Get-QADUser $u.manager).DisplayName}else{$label2.Hide()}
                
                    $U_made.Text = $($u.whenCreated.ToString('dd/MM/yy HH:mm:ss'))
                    $U_changed.Text = $($u.whenChanged.ToString('dd/MM/yy HH:mm:ss'))
                    $Uid_name.Text = $($u.CanonicalName)

                    if ($u.PasswordLastSet){ 
                        $pass_made.show()
                        $pass_made.text = $($u.PasswordLastSet.ToString('dd/MM/yy HH:mm:ss')) 
                        } else { $pass_made.Hide() }
                    $pass_life.text = '{0:N3}' -f $($u.PasswordAge.TotalDays)
                    
                    if ($u.PasswordIsExpired -eq $true){
                        $pass_stat.visible = $true
                        $pass_stat.text = "Пароль просрочен!"
                        $pass_stat.forecolor = "red"
                        }
                    if ($u.UserMustChangePassword -eq $true){
                        $pass_stat.visible = $true
                        $pass_stat.text = "Требуется сменить пароль при входе!"
                        $pass_stat.forecolor = "orange"
                        }
                    #$pass_stat.text = $($u.PasswordStatus)

                    if ($u.AccountIsLockedOut -eq $true){
                        $U_stat.visible = $true
                        $U_stat.text = "Аккаунт залочен (LockedOut)!"
                        $U_stat.forecolor = "orange"
                        }
                    if ($u.AccountIsDisabled -eq $true){
                        $U_stat.visible = $true
                        $U_stat.text = "Аккаунт отключен (Disabled)!"
                        $U_stat.forecolor = "red"
                        }
                    
                    $GroupsListBox.Items.Clear()
                    $Groups = $($u.memberof).split(",") | ForEach-Object {if($_.StartsWith('CN=')){$_.remove(0,3)}}
                    #$Groups = foreach ($Group in $GroupsSplit){if($Group.StartsWith('CN=')){$Group.remove(0,3)}}
                    if($Groups){$GroupsListBox.Items.AddRange($Groups);}
                    if(($Groups -contains 'THFarm_users') -or ($Groups -contains 'VTSFarm_users') -or ($Groups -contains 'VTSFarm_users_disk1') -or ($Groups -contains 'VTSFarm_users_disk2') -or ($Groups -contains 'VTSFarm_users_disk3')){ $TermsProfile_But.Enabled = $true }                    
                    
                    #"Нужно сменить пароль: $($u.UserMustChangePassword)"
                    #"`n Пользователь был замечен на следующих локальных компьютерах:"
                    #вывод информации из SCOM
                    #$SCCMout
                    #"`n"
            }
        }
    }

    function Show_Alert($var){
            $Alert_Window = New-Object System.Windows.Forms.Form                        
            $Alert = New-Object System.Windows.Forms.Label;
            $AlertButOk = New-Object System.Windows.Forms.Button;
            $AlertButCancel = New-Object System.Windows.Forms.Button;
                                        
            $Alert.TextAlign = "MiddleCenter"

            $AlertButOk.Anchor = "Bottom";
			$AlertButOk.AutoEllipsis = $true;
			$AlertButOk.Cursor = "Hand";
			$AlertButOk.Location = New-Object System.Drawing.Point(67, 126);			
			$AlertButOk.Size = New-Object System.Drawing.Size(72, 24);
            $AlertButOk.Text = "Окай";
			$AlertButOk.UseVisualStyleBackColor = $true;

            $AlertButCancel.Anchor = "Bottom";
			$AlertButCancel.AutoEllipsis = $true;
			$AlertButCancel.Cursor = "Hand";
			$AlertButCancel.Location = New-Object System.Drawing.Point(143, 126);			
			$AlertButCancel.Size = New-Object System.Drawing.Size(72, 24);
            $AlertButCancel.Text = "Нееет";
			$AlertButCancel.UseVisualStyleBackColor = $true;                    
            $AlertButCancel.add_Click({$Alert_Window.Close()})

            switch ($var){
                0 { 
                    $Alert.Location = New-Object System.Drawing.Point(12, 40);
                    $Alert.Size = New-Object System.Drawing.Size(260, 42);
                    $Alert.text = "Не найдена оснастка Quest.ActiveRoles.ADManagement.`nПерейти на страницу загрузки оснастки?"
                  
                    $AlertButOk.add_Click({
                        $ie = new-object -com "InternetExplorer.Application"
                        $ie.navigate("http://www.quest.com/powershell/activeroles-server.aspx")
                        $ie.visible = $true
                        $Alert_Window.Close()
                        $Window.Close()})
                    }
                1 { 
                    $Alert.Location = New-Object System.Drawing.Point(12, 40);
                    $Alert.Size = New-Object System.Drawing.Size(260, 42);
                    $Alert.text = "Пользователь " + $UserSearch_Box.Text +" не найден!"
                  
                    $AlertButOk.Location = New-Object System.Drawing.Point(106, 126);
                    $AlertButCancel.Hide()
                    $AlertButOk.add_Click({$Alert_Window.Close()})
                    }
                2 { 
                    $Alert.Location = New-Object System.Drawing.Point(12, 9);
                    $Alert.Size = New-Object System.Drawing.Size(260, 20);
                    $Alert.text = "Найдено несколько пользователей, уточните:`n"                   
                    $AlertListBox = New-Object System.Windows.Forms.ListBox;
                    $AlertListBox.Location = New-Object System.Drawing.Point(12, 32);			
			        $AlertListBox.Size = New-Object System.Drawing.Size(260, 82);
                    $Alert_Window.Controls.Add($AlertListBox);
                    $AlertListBox.Anchor = "Top"
                    $AlertListBox.Items.AddRange($user)
                     
                    $AlertListBox.SelectedIndex = 0                                       
                    $AlertListBox.add_keyDown({if ($_.KeyCode -eq 'Return'){ $Alert_Window.Close(); $UserSearch_Box.text = $AlertlistBox.SelectedItem; Search_User }})                  
                    $AlertButOk.add_Click({ $Alert_Window.Close(); $UserSearch_Box.text = $AlertlistBox.SelectedItem; Search_User })
                    }   
                 3 { 
                    $Alert.Location = New-Object System.Drawing.Point(12, 40);
                    $Alert.Size = New-Object System.Drawing.Size(260, 42);                   
                    $Alert.text = "Сбросить пароль у $user ?"                    
                  
                    $AlertButOk.add_Click({
                        Unlock-QADUser $user | Out-Null
                        Set-QADUser $user -UserMustChangePassword $true -UserPassword 'Pa$$123456' | Out-Null
                        $Alert_window.close()
                        })
                    }
                  4 {
                     $Alert.Location = New-Object System.Drawing.Point(12, 5)
                     $Alert.Size = New-Object System.Drawing.Size(260, 20)
                     $Alert.Text = "Сменить пароль для $user"

                     $label6 = New-Object System.Windows.Forms.Label
                     $label7 = New-Object System.Windows.Forms.Label
                     $label6.Location = New-Object System.Drawing.Point(12, 30)
                     $label7.Location = New-Object System.Drawing.Point(12, 60)
                     $label6.Size = New-Object System.Drawing.Size(100, 20)
                     $label7.Size = New-Object System.Drawing.Size(100, 20)
                     $label6.Text = "Новый пароль:"
                     $label7.Text = "Подтверждение:"
                     $label6.TextAlign = "MiddleLeft" 
                     $label7.TextAlign = "MiddleLeft" 

                     $MaskNew = New-Object System.Windows.Forms.MaskedTextBox
                     $MaskConf = New-Object System.Windows.Forms.MaskedTextBox
                     $MaskNew.Location = New-Object System.Drawing.Point(112, 30)
                     $MaskConf.Location = New-Object System.Drawing.Point(112, 60)
                     $MaskNew.Size = New-Object System.Drawing.Size(154, 20)
                     $MaskConf.Size = New-Object System.Drawing.Size(154, 20)
                     $MaskNew.PasswordChar = '*'
                     $MaskConf.PasswordChar = '*'

                     $PassCheckBox = New-Object System.Windows.Forms.CheckBox
                     $PassCheckBox.Location = New-Object System.Drawing.Point(40, 90)
                     $PassCheckBox.Size = New-Object System.Drawing.Point(210, 24)
                     $PassCheckBox.Text = "Должен сменить пароль при входе"

                     $AlertButOk.Enabled = $false
                     $MaskConf.add_keyUp({if ($MaskNew.text -ceq $MaskConf.text){ $AlertButOk.Enabled = $true}else{$AlertButOk.Enabled = $false}})
                     $PassCheckBox.add_CheckedChanged({$AlertButOk.Enabled = $true})
                     $AlertButOk.add_Click({
                        Unlock-QADUser $user | Out-Null
                        if ($PassCheckBox.Checked -eq $true){Set-QADUser $user -UserMustChangePassword $true | Out-Null}                       
                        if ($MaskNew.Text){Set-QADUser $user -UserPassword $($MaskNew.Text) | Out-Null}
                        $Alert_window.Close()
                        })

                    }
                5 { 
                    $Alert.Location = New-Object System.Drawing.Point(12, 40);
                    $Alert.Size = New-Object System.Drawing.Size(260, 42);
                    $Alert.text = "Сбросить сессию $user ($sessnum на $cursrv)?"
                    $allow = $false
                    $AlertButOk.add_Click({$allow = $true; $AlertButOk.enabled = $false; Logoff_User; $Alert_Window.Close()})
                   }
             }
                                


            $Alert.SuspendLayout();
            $Alert_Window.SuspendLayout();          
            #$Alert_Window.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13);
			$Alert_Window.AutoScaleMode = "Font";
			$Alert_Window.AutoScroll = $true;
			$Alert_Window.Controls.Add($Alert);
            $Alert_Window.Controls.Add($AlertButOk);            
            $Alert_Window.Controls.Add($AlertButCancel);
            $Alert_Window.Controls.Add($label6);
            $Alert_Window.Controls.Add($label7) 
            $Alert_Window.Controls.Add($MaskNew)
            $Alert_Window.Controls.Add($MaskConf)
            $Alert_Window.Controls.Add($PassCheckBox)
			$Alert_Window.MinimumSize = New-Object System.Drawing.Size (300, 200);
            $Alert_Window.MaximumSize = New-Object System.Drawing.Size (300, 200);
			$Alert_Window.StartPosition = "CenterScreen";
            $Alert_Window.SizeGripStyle = "Hide";
			$Alert_Window.Text = "Внимание!";
            $Alert_Window.Topmost = $true;            

            $Alert_window.ShowDialog() | Out-Null
            }
#$global:cursrv = $null
#$global:sessnum = $null
#$global:status = $null
function Find_Terms(){
    $terms = $null
	
    if ($UseTH.Checked -eq $true){	
	    if (Get-QADUser $user -indirectMemberOf THFarm_users){
        $terms = get-qadcomputer "th-*" | where { ($_.userAccountControl -ne 4098) -and ($_.name -ne 'TH-TEST') -and ($_.name -ne 'TH-BROKER')} | select-object name | % {$_.name.toString()}         
        } else { Show_Alert 1 }
    }
     if ($UseVTS.Checked){
        if ((Get-QADUser $user -indirectMemberOf VTSFarm_users) -or (Get-QADUser $user -indirectMemberOf VTSFarm_users_disk1) -or (Get-QADUser $user -indirectMemberOf VTSFarm_users_disk2) -or (Get-QADUser $user -indirectMemberOf VTSFarm_users_disk3)) {
        $terms += get-qadcomputer "vts-*" | where { ($_.userAccountControl -ne 4098) -and ($_.name -ne 'VTS-DATA') -and ($_.name -ne 'VTS-BROKER')} | select-object name | % {$_.name.toString()} 
        } else { Show_Alert 1 }
    }

    $FindTerms_But.Enabled = $false
    $TermsList.Hide()
    $PBarTerms.Show()
    $PBarStat.Show()
    $TermsLogoff_But.Enabled = $false
    $TermsRA_But.Enabled = $false
    $PBarTerms.Value = 0
    $PBarTerms.Maximum = $terms.count
    $TermsList.Text = $null
     
    $terms | sort-object -unique | % {    
    $srv = $_

    $PBarStat.Text = "$srv"
    if (test-connection $srv -count 1 -quiet) {
        $session_tmp = @(query user $user /server:$srv)
        if($session_tmp[1] -like "* $user *"){            

            $session = $session_tmp[1].trim()                
            while ($session.indexOF("  ") -ne -1) {$session = $session.Replace("  "," ")}
            $session = $session.Split(" ")
            [System.Object[]]$sesinfo = $null
            if ($session.Count -eq 6) { $i=1 } 
            if ($session.Count -eq 7) { $i=2 }
            if ($i) { 
                    $global:cursrv = $srv
                    $global:sessnum = $session[$i]
                    $global:status = $i - 1
                    $sesinfo += "Сервер: $cursrv"
                    $sesinfo += "Логин: $($session[0])"
                    $sesinfo += "Сессия: $sessnum"
                    $sesinfo += "Дата входа: $($session[$i+3])"
                    $sesinfo += "Время входа: $($session[$i+4])"
                    $sesinfo += "Статус: $status"                    
              
                }
                $TermsList.Show()
                $TermsList.Lines += $sesinfo + ' ' +$sessnum + ' ' +$cursrv

                $TermsLogoff_But.Enabled = $true               
                $TermsRA_But.Enabled = $true
            }            
        }
         $PBarTerms.Increment(1)
    }
    $FindTerms_But.Enabled = $True
    $PBarTerms.Hide()
    $PBarStat.Hide()
}

function Logoff_User(){
    
    if($cursrv -and ($allow -eq $true)){
            "Завершаем сессию $user ($sessnum) на сервере $cursrv..."
            Logoff $sessnum /server:$cursrv
            }

    }

function StartRA(){


    }

$UseTH.add_CheckedChanged({if ($UseTH.Checked -or $UseVTS.Checked){$FindTerms_But.enabled = $true}else{$FindTerms_But.enabled = $false}})
$UseVTS.add_CheckedChanged({if ($UseVTS.Checked -or $UseTH.Checked){$FindTerms_But.enabled = $true}else{$FindTerms_But.enabled = $false}})

$TermsRA_But.add_Click({ StartRA })
$TermsRA_But.add_keyDown({if ($_.KeyCode -eq 'Return'){ StartRA }})

$TermsLogoff_But.add_Click({ Show_Alert 5 })
$TermsLogoff_But.add_keyDown({if ($_.KeyCode -eq 'Return'){ Show_Alert 5 }})

$FindTerms_But.add_Click({ Find_Terms })
$FindTerms_But.add_keyDown({if ($_.KeyCode -eq 'Return'){ Find_Terms }})
                    
$UserSearch_But.add_Click({ Search_User })
$UserSearch_Box.add_keyDown({if ($_.KeyCode -eq 'Return'){ Search_User }})

$PassReset_But.add_Click({ Show_alert 3 })
$PassReset_But.add_keyDown({if ($_.KeyCode -eq 'Return'){ Show_Alert 3 }})

$PassChange_But.add_Click({ Show_Alert 4 })
$PassChange_But.add_keyDown({if ($_.KeyCode -eq 'Return'){ Show_Alert 4 }})


$Window.add_Shown({ $usersearch_box.Focus() })
$Window.ShowDialog() | Out-Null
Disconnect-QADService
