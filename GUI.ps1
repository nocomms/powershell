Add-Type -assembly System.Windows.Forms | Out-Null


$Window = New-Object System.Windows.Forms.Form
$Window.SuspendLayout();

            $MainBox = New-Object System.Windows.Forms.TabControl;
            $MainBox.SuspendLayout();

            $UserBox = New-Object System.Windows.Forms.TabPage;
            $_tab = New-Object System.Windows.Forms.TabPage;
$_tab.SuspendLayout();
            $label1 = New-Object System.Windows.Forms.Label;
            $UserSearch_Box = New-Object System.Windows.Forms.TextBox;
            $UserSearch_But = New-Object System.Windows.Forms.Button;

            $UserInfo = New-Object System.Windows.Forms.GroupBox;
            $U_stat = New-Object System.Windows.Forms.Label;
			$U_fullname = New-Object System.Windows.Forms.Label;
            $U_login = New-Object System.Windows.Forms.Label;
            $U_id = New-Object System.Windows.Forms.Label;
            $U_email = New-Object System.Windows.Forms.Label;
            $U_phones = New-Object System.Windows.Forms.Label;
            $label2 = New-Object System.Windows.Forms.Label;
            $LinkToBoss = New-Object System.Windows.Forms.LinkLabel;
            $label3 = New-Object System.Windows.Forms.Label;
            $U_made = New-Object System.Windows.Forms.Label;
            $label4 = New-Object System.Windows.Forms.Label;
            $U_changed = New-Object System.Windows.Forms.Label;
            $Uid_name = New-Object System.Windows.Forms.Label;
            $UserBox.SuspendLayout();
            $UserInfo.SuspendLayout();

            $PassInfo = New-Object System.Windows.Forms.GroupBox;
			$Pass_stat = New-Object System.Windows.Forms.Label;
            $label5 = New-Object System.Windows.Forms.Label;
            $pass_made = New-Object System.Windows.Forms.Label;
			$pass_life = New-Object System.Windows.Forms.Label;
			$PassReset_But = New-Object System.Windows.Forms.Button;
            $PassChange_But = New-Object System.Windows.Forms.Button;
			$LockOutStat_But = New-Object System.Windows.Forms.Button;
			$PassInfo.SuspendLayout();

	############################################
  
            ## 
			## UserBox
			## 
			$UserBox.BackColor = "white";
			$UserBox.Controls.Add($PassInfo);
			$UserBox.Controls.Add($UserInfo);
			$UserBox.Controls.Add($UserSearch_But);
			$UserBox.Controls.Add($UserSearch_Box);
			$UserBox.Controls.Add($label1);
			$UserBox.Location = New-Object System.Drawing.Point(4, 22);
			$UserBox.Name = "UserBox";
			$UserBox.Padding = New-Object System.Windows.Forms.Padding(3);
			$UserBox.Size = New-Object System.Drawing.Size(976, 679);
			$UserBox.TabIndex = 0;
			$UserBox.Text = "Карточка пользователя";   
                ## 
			## tab
			## 
			$tab.BackColor = "white";
			$tab.Location = New-Object System.Drawing.Point(4, 22);
			$tab.Name = "tab";
			$tab.Padding = New-Object System.Windows.Forms.Padding(3);
			$tab.Size = New-Object System.Drawing.Size(976, 679);
			$tab.TabIndex = 0;
			$tab.Text = "tab";          
			## 
			## PassInfo
			## 
			$PassInfo.BackColor = "Whitesmoke";
			$PassInfo.Controls.Add($Pass_stat);
			$PassInfo.Controls.Add($pass_life);
			$PassInfo.Controls.Add($pass_made);
			$PassInfo.Controls.Add($LockOutStat_But);
			$PassInfo.Controls.Add($PassChange_But);
			$PassInfo.Controls.Add($PassReset_But);
			$PassInfo.Controls.Add($label5);
			$PassInfo.Location = New-Object System.Drawing.Point(15, 289);
			$PassInfo.Name = "PassInfo";
			$PassInfo.Size = New-Object System.Drawing.Size(400, 80);
			$PassInfo.TabIndex = 4;
			$PassInfo.TabStop = $false;
			$PassInfo.Text = "Пароль";
			## 
			## Pass_stat
			## 
			$Pass_stat.Anchor = @("Top","Right");
			$Pass_stat.BackColor = "Transparent";
			$Pass_stat.Location = New-Object System.Drawing.Point(250, 5);
			$Pass_stat.Name = "Pass_stat";
			$Pass_stat.Size = New-Object System.Drawing.Size(150, 20);
			$Pass_stat.TabIndex = 0;
			$Pass_stat.Text = "(pass_stat)";
			$Pass_stat.TextAlign = "MiddleRight";
			$Pass_stat.Visible = $false;
			## 
			## pass_life
			## 
			$pass_life.Location = New-Object System.Drawing.Point(250, 25);
			$pass_life.Name = "pass_life";
			$pass_life.Size = New-Object System.Drawing.Size(102, 20);
			$pass_life.TabIndex = 0;
			$pass_life.Text = "(pass_life)";
			$pass_life.TextAlign = "MiddleLeft";
			## 
			## pass_made
			## 
			$pass_made.Location = New-Object System.Drawing.Point(86, 25);
			$pass_made.Name = "pass_made";
			$pass_made.Size = New-Object System.Drawing.Size(158, 20);
			$pass_made.TabIndex = 0;
			$pass_made.Text = "(pass_made)";
			$pass_made.TextAlign = "MiddleLeft";
			## 
			## LockOutStat_But
			## 
			$LockOutStat_But.Location = New-Object System.Drawing.Point(240, 50);
			$LockOutStat_But.Name = "LockOutStat_But";
			$LockOutStat_But.Size = New-Object System.Drawing.Size(80, 24);
			$LockOutStat_But.TabIndex = 3;
			$LockOutStat_But.Text = "LockStat";
			$LockOutStat_But.UseVisualStyleBackColor = $true;
			## 
			## PassChange_But
			## 
			$PassChange_But.Location = New-Object System.Drawing.Point(154, 50);
			$PassChange_But.Name = "PassChange_But";
			$PassChange_But.Size = New-Object System.Drawing.Size(80, 24);
			$PassChange_But.TabIndex = 2;
			$PassChange_But.Text = "Изменить";
			$PassChange_But.UseVisualStyleBackColor = $true;
			## 
			## PassReset_But
			## 
			$PassReset_But.Location = New-Object System.Drawing.Point(68, 50);
			$PassReset_But.Name = "PassReset_But";
			$PassReset_But.Size = New-Object System.Drawing.Size(80, 24);
			$PassReset_But.TabIndex = 1;
			$PassReset_But.Text = "Сбросить";
			$PassReset_But.UseVisualStyleBackColor = $true;
			## 
			## label5
			## 
			$label5.Location = New-Object System.Drawing.Point(6, 25);
			$label5.Name = "label5";
			$label5.Size = New-Object System.Drawing.Size(80, 20);
			$label5.TabIndex = 0;
			$label5.Text = "Установлен:";
			$label5.TextAlign = "MiddleLeft";
			## 
			## UserInfo
			## 
			$UserInfo.BackColor = "Whitesmoke";
			$UserInfo.Controls.Add($Uid_name);
			$UserInfo.Controls.Add($label4);
			$UserInfo.Controls.Add($label3);
			$UserInfo.Controls.Add($U_changed);
			$UserInfo.Controls.Add($U_made);
			$UserInfo.Controls.Add($U_fullname);
			$UserInfo.Controls.Add($U_stat);
			$UserInfo.Controls.Add($LinkToBoss);
			$UserInfo.Controls.Add($U_phones);
			$UserInfo.Controls.Add($label2);
			$UserInfo.Controls.Add($U_email);
			$UserInfo.Controls.Add($U_login);
			$UserInfo.Controls.Add($U_id);
			$UserInfo.Location = New-Object System.Drawing.Point(15, 43);
			$UserInfo.Name = "UserInfo";
			$UserInfo.Size = New-Object System.Drawing.Size(400, 240);
			$UserInfo.TabIndex = 3;
			$UserInfo.TabStop = $false;
			$UserInfo.Text = "Общее";
			## 
			## Uid_name
			## 
			$Uid_name.Location = New-Object System.Drawing.Point(6, 209);
			$Uid_name.Name = "Uid_name";
			$Uid_name.Size = New-Object System.Drawing.Size(360, 20);
			$Uid_name.TabIndex = 0;
			$Uid_name.Text = "(uid_name)";
			$Uid_name.TextAlign = "MiddleLeft";
			## 
			## label4
			## 
			$label4.Location = New-Object System.Drawing.Point(6, 189);
			$label4.Name = "label4";
			$label4.Size = New-Object System.Drawing.Size(88, 20);
			$label4.TabIndex = 0;
			$label4.Text = "Изменён:";
			$label4.TextAlign = "MiddleLeft";
			## 
			## label3
			## 
			$label3.Location = New-Object System.Drawing.Point(6, 169);
			$label3.Name = "label3";
			$label3.Size = New-Object System.Drawing.Size(88, 20);
			$label3.TabIndex = 0;
			$label3.Text = "Создан:";
			$label3.TextAlign = "MiddleLeft";
			## 
			## U_changed
			## 
			$U_changed.Location = New-Object System.Drawing.Point(100, 189);
			$U_changed.Name = "U_changed";
			$U_changed.Size = New-Object System.Drawing.Size(266, 20);
			$U_changed.TabIndex = 0;
			$U_changed.Text = "(u_change)";
			$U_changed.TextAlign = "MiddleLeft";
			## 
			## U_made
			## 
			$U_made.Location = New-Object System.Drawing.Point(100, 169);
			$U_made.Name = "U_made";
			$U_made.Size = New-Object System.Drawing.Size(266, 20);
			$U_made.TabIndex = 0;
			$U_made.Text = "(u_made)";
			$U_made.TextAlign = "MiddleLeft";
			## 
			## U_fullname
			## 
			$U_fullname.Location = New-Object System.Drawing.Point(6, 25);
			$U_fullname.Name = "U_fullname";
			$U_fullname.Size = New-Object System.Drawing.Size(360, 20);
			$U_fullname.TabIndex = 0;
			$U_fullname.Text = "(fullname)";
			$U_fullname.TextAlign = "MiddleLeft";
			## 
			## U_stat
			## 
			$U_stat.Anchor = @("Top","Right");
			$U_stat.BackColor = "Transparent";#System.Drawing.Color.Transparent;
			$U_stat.Location = New-Object System.Drawing.Point(250, 5);
			$U_stat.Name = "U_stat";
			$U_stat.Size = New-Object System.Drawing.Size(150, 20);
			$U_stat.TabIndex = 0;
			$U_stat.Text = "(status)";
			$U_stat.TextAlign = "MiddleRight";#System.Drawing.ContentAlignment.MiddleRight;
			$U_stat.Visible = $false;
			## 
			## LinkToBoss
			## 
			$LinkToBoss.Location = New-Object System.Drawing.Point(86, 128);
			$LinkToBoss.Name = "LinkToBoss";
			$LinkToBoss.Size = New-Object System.Drawing.Size(280, 20);
			$LinkToBoss.TabIndex = 1;
			$LinkToBoss.TabStop = $true;
			$LinkToBoss.Text = "(boss)";
			$LinkToBoss.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			$LinkToBoss.VisitedLinkColor = "Blue";#System.Drawing.Color.Blue;
			## 
			## U_phones
			## 
			$U_phones.Location = New-Object System.Drawing.Point(6, 105);
			$U_phones.Name = "U_phones";
			$U_phones.Size = New-Object System.Drawing.Size(360, 20);
			$U_phones.TabIndex = 0;
			$U_phones.Text = "(phones)";
			$U_phones.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			## 
			## label2
			## 
			$label2.Location = New-Object System.Drawing.Point(6, 128);
			$label2.Name = "label2";
			$label2.Size = New-Object System.Drawing.Size(85, 20);
			$label2.TabIndex = 0;
			$label2.Text = "Руководитель:";
			$label2.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			## 
			## U_email
			## 
			$U_email.Location = New-Object System.Drawing.Point(6, 85);
			$U_email.Name = "U_email";
			$U_email.Size = New-Object System.Drawing.Size(360, 20);
			$U_email.TabIndex = 0;
			$U_email.Text = "(email)";
			$U_email.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			## 
			## U_login
			## 
			$U_login.Location = New-Object System.Drawing.Point(6, 65);
			$U_login.Name = "U_login";
			$U_login.Size = New-Object System.Drawing.Size(360, 20);
			$U_login.TabIndex = 0;
			$U_login.Text = "(login)";
			$U_login.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			## 
			## U_id
			## 
			$U_id.Location = New-Object System.Drawing.Point(6, 45);
			$U_id.Name = "U_id";
			$U_id.Size = New-Object System.Drawing.Size(360, 20);
			$U_id.TabIndex = 0;
			$U_id.Text = "(id)";
			$U_id.TextAlign = "MiddleLeft";#System.Drawing.ContentAlignment.MiddleLeft;
			## 
			## UserSearch_But
			## 
			$UserSearch_But.Anchor = @("Left", "Top");
			$UserSearch_But.AutoEllipsis = $true;
			$UserSearch_But.Cursor = "Hand";
			$UserSearch_But.Location = New-Object System.Drawing.Point(387, 13);
			$UserSearch_But.Name = "UserSearch_But";
			$UserSearch_But.Size = New-Object System.Drawing.Size(24, 24);
			$UserSearch_But.TabIndex = 2;
			$UserSearch_But.Text = "⚲";
			$UserSearch_But.UseVisualStyleBackColor = $true;
			## 
			## UserSearch_Box
			## 
			$UserSearch_Box.Location = New-Object System.Drawing.Point(131, 16);
			$UserSearch_Box.Name = "UserSearch_Box";
			$UserSearch_Box.Size = New-Object System.Drawing.Size(250, 20);
			$UserSearch_Box.TabIndex = 1;
            $UserSearch_Box.MaxLength = 24
			## 
			## label1
			## 
			$label1.Location = New-Object System.Drawing.Point(15, 15);
			$label1.Name = "label1";
			$label1.Size = New-Object System.Drawing.Size(110, 20);
			$label1.TabIndex = 0;
			$label1.Text = "Имя пользователя:";
			$label1.TextAlign = "MiddleCenter";#System.Drawing.ContentAlignment.MiddleCenter;
			## 
			## MainBox
			## 
			$MainBox.Anchor = @("Top", "Right", "Left", "Bottom");
			$MainBox.Controls.Add($UserBox);
			$MainBox.Location = New-Object System.Drawing.Point(12, 12);
			$MainBox.Name = "MainBox";
			$MainBox.SelectedIndex = 0;
			$MainBox.Size = New-Object System.Drawing.Size(984, 705);
		###$MainBox.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
			$MainBox.TabIndex = 0;
			## 
			## MainForm
			## 
			$Window.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13);
			$Window.AutoScaleMode = "Font";
			$Window.AutoScroll = $true;
			$Window.ClientSize = New-Object System.Drawing.Size (1008, 729);
			$Window.Controls.Add($MainBox);
			$Window.MinimumSize = New-Object System.Drawing.Size (1024, 768);
            $Window.MaximumSize = New-Object System.Drawing.Size (1024, 768);
			$Window.Name = "MainForm";
			$window.StartPosition = "CenterScreen";
            $Window.SizeGripStyle = "Hide";
			$Window.Text = "HelpDesk GUI Tool";
            $tab.ResumeLayout($false);
			$tab.PerformLayout();
			$UserBox.ResumeLayout($false);
			$UserBox.PerformLayout();
			$PassInfo.ResumeLayout($false);
			$UserInfo.ResumeLayout($false);
			$MainBox.ResumeLayout($false);
			$Window.ResumeLayout($false);

function Search_User(){
    if ($UserSearch_Box.Text -ne ''){
        $data = $UserSearch_Box.Text
        
        $LinkToBoss.text = $data
        $pass_stat.visible = $true
        }
        }
#        Connect-QADService -service 'hq-dc.okmarket.ru' | Out-Null
#        while ($data){
#            $u = get-qaduser $username 
#            $user = $u.SamAccountName
#            if (!$user) {
#                "`nПользователь $username не найден!"
#                $variant = -1
#                } else
#                if (($user.gettype()).name -eq 'String') {
#                    "`nНайден пользователь: $user"
#   	                $SCCMout = $null
#  	                $SQLCon = New-Object System.Data.SQLClient.SQLConnection
#	                $SQLSrv = 'central-sccm'
#	                $SQLCon.ConnectionString = "Server=$SQLSrv; Integrated Security = True"
#	                $SQLCon.Open()
#	                $Query='SELECT * FROM [CM_MCM].[dbo].[v_SUP_ConsolUser] where "console use" ='+"'okey\$username'"
#	                $SQLCommand = New-Object System.Data.SqlClient.SqlCommand($Query, $SQLCon)
#	                $SQLAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($SQLCommand)
#	                $DataSet = New-Object System.Data.DataSet
#	                $SQLAdapter.Fill($DataSet) >> $null
#	                $SCCMout = $DataSet.Tables[0] | select "Name PC","Last Console Use" | Sort-Object -Property "Last Console Use" -Descending | where{$_."Name PC" -notlike "vts-*"}
#	                $SQLCon.Close()
#                   }


$UserSearch_But.add_Click({ Invoke-Expression "Search_User" })

$Window.ShowDialog() | Out-Null
