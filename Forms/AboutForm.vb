Imports System.Drawing
Imports System.Windows.Forms


    Public Class AboutForm
        Inherits Form

        Private WithEvents btnOK As Button
        Private WithEvents lnkFB As LinkLabel
        Private WithEvents lnkAnthropic As LinkLabel

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "About " & APP_NAME
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Size = New Size(440, 380)
            Me.Font = New Font("Segoe UI", 9)

            ' Header banner
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 75,
                .BackColor = Color.FromArgb(50, 50, 100)
            }

            Dim lblTitle As New Label() With {
                .Text = APP_NAME,
                .Font = New Font("Segoe UI", 20, FontStyle.Bold),
                .ForeColor = Color.White,
                .BackColor = Color.Transparent,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Top,
                .Height = 45
            }

            Dim lblSubtitle As New Label() With {
                .Text = "A FreeBASIC Development Environment",
                .Font = New Font("Segoe UI", 9, FontStyle.Italic),
                .ForeColor = Color.FromArgb(220, 220, 220),
                .BackColor = Color.Transparent,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Top,
                .Height = 25
            }

            pnlHeader.Controls.Add(lblSubtitle)
            pnlHeader.Controls.Add(lblTitle)

            ' Info labels
            Dim y = 90
            Dim lf = 20

            Dim lblVersion As New Label() With {
                .Text = "Version " & APP_VERSION,
                .Font = New Font("Segoe UI", 11),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, 24),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += 35

            Dim lblAuthor As New Label() With {
                .Text = "Created by Ronen Blumberg",
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, lf),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += lf + 4

            Dim lblCopyright As New Label() With {
                .Text = "Copyright © 2026 Ronen Blumberg. All rights reserved.",
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, lf),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += lf + 12

            Dim lblDesc As New Label() With {
                .Text = "FBEditor is a modern FreeBASIC IDE with Scintilla editor," & vbCrLf &
                        "AI-powered coding assistance (Claude), full build system," & vbCrLf &
                        "UTF-8/ANSI encoding, dark/light themes, and code outline.",
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(10, y),
                .Size = New Size(Me.ClientSize.Width - 20, 56),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += 64

            Dim lblBuiltWith As New Label() With {
                .Text = "Built with VB.NET / .NET Framework 4.8 / ScintillaNET",
                .TextAlign = ContentAlignment.MiddleCenter,
                .ForeColor = Color.Gray,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, lf),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += lf + 8

            ' Links
            lnkFB = New LinkLabel() With {
                .Text = "FreeBASIC: www.freebasic.net",
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, lf),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += lf + 2

            lnkAnthropic = New LinkLabel() With {
                .Text = "AI powered by Claude (Anthropic)",
                .TextAlign = ContentAlignment.MiddleCenter,
                .Location = New Point(0, y),
                .Size = New Size(Me.ClientSize.Width, lf),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            }
            y += lf + 16

            ' OK button
            btnOK = New Button() With {
                .Text = "OK",
                .Size = New Size(100, 32),
                .Location = New Point((Me.ClientSize.Width - 100) \ 2, y),
                .DialogResult = DialogResult.OK,
                .Anchor = AnchorStyles.Top
            }

            Me.AcceptButton = btnOK
            Me.CancelButton = btnOK

            Me.Controls.AddRange({pnlHeader, lblVersion, lblAuthor, lblCopyright, lblDesc,
                                  lblBuiltWith, lnkFB, lnkAnthropic, btnOK})
        End Sub

        Private Sub LnkFB_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkFB.LinkClicked
            SafeProcessStart("https://www.freebasic.net")
        End Sub

        Private Sub LnkAnthropic_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkAnthropic.LinkClicked
            SafeProcessStart("https://www.anthropic.com")
        End Sub

    End Class

