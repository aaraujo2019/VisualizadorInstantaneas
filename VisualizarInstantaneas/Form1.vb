Option Explicit On
Imports System.Configuration
Imports System.Data.SqlClient

Public Class FrmInstantaneas
    Private TiempoRestante As Integer
    Private dt As DataTable
    Dim Da As New SqlDataAdapter
    Dim Cmd As New SqlCommand
    Dim Dataset As New DataTable
    Dim Cn As New SqlConnection(ConfigurationSettings.AppSettings("StringConection"))
    Dim cnStr As String
    Dim control As Integer

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = Screen.PrimaryScreen.WorkingArea.Location
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size
        'control = Convert.ToInt32(ConfigurationSettings.AppSettings("ControlSeleccionado"))
        control = 1
        'cargarinstantaneas()
        Call TiempoEjecutar(15)
    End Sub



    Private Sub cargarinstantaneas()
        Try
            If My.Computer.Network.Ping("SEGSVRSQL01") Then
                LblServer.Text = "ON-LINE"
                LblServer.BackColor = Color.White

                OcultarCampos()
                Lbltipo2.Text = "Tenor"
                LabelTipo.Text = "Ppm"
                Lbltipo2.BackColor = Color.White

                Da = New SqlDataAdapter("SELECT UltimaFecha, HoraActualizacion, ubicacion, Tenor_ppm FROM dbo.PB_InstantaneasVisualizadorFinal order by ubicacion", Cn)
                Da.Fill(Dataset)

                If Dataset.Rows.Count > 0 Then
                    Dim registro = 1
                    For Each ResumenTabla As DataRow In Dataset.Rows
                        If registro = 1 Then
                            Fecha1.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora1.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion1.Text = CStr(ResumenTabla("ubicacion"))
                            tenor1.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 2 Then
                            Fecha2.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora2.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion2.Text = CStr(ResumenTabla("ubicacion"))
                            tenor2.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 3 Then
                            Fecha3.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora3.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion3.Text = CStr(ResumenTabla("ubicacion"))
                            tenor3.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 4 Then
                            Fecha4.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora4.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion4.Text = CStr(ResumenTabla("ubicacion"))
                            tenor4.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 5 Then
                            Fecha5.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora5.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion5.Text = CStr(ResumenTabla("ubicacion"))
                            tenor5.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 6 Then
                            Fecha6.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora6.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion6.Text = CStr(ResumenTabla("ubicacion"))
                            tenor6.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 7 Then
                            Fecha7.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora7.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion7.Text = CStr(ResumenTabla("ubicacion"))
                            tenor7.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 8 Then
                            Fecha8.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora8.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion8.Text = CStr(ResumenTabla("ubicacion"))
                            tenor8.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 9 Then
                            Fecha9.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora9.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion9.Text = CStr(ResumenTabla("ubicacion"))
                            tenor9.Text = CStr(ResumenTabla("Tenor_ppm"))
                        ElseIf registro = 10 Then
                            Fecha10.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora10.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion10.Text = CStr(ResumenTabla("ubicacion"))
                            tenor10.Text = CStr(ResumenTabla("Tenor_ppm"))

                        End If
                        registro = registro + 1
                    Next
                Else
                    LblServer.BackColor = Color.Red
                End If
            Else
                LblServer.Text = "OFF-LINE"
                LblServer.BackColor = Color.Red
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cargarDensidades()
        Try
            If My.Computer.Network.Ping("SEGSVRSQL01") Then
                LblServer.Text = "ON-LINE"
                LblServer.BackColor = Color.White

                OcultarCampos()

                Lbltipo2.Text = "Densidades"
                LabelTipo.Text = "Densidades"
                Lbltipo2.BackColor = Color.White

                Dim registro = 0

                Da = New SqlDataAdapter("SELECT UltimaFecha, HoraActualizacion, ubicacion, Densidad FROM PB_DensidadesVisualizador order by ubicacion", Cn)
                Da.Fill(Dataset)

                If Dataset.Rows.Count > 0 Then
                    For Each ResumenTabla As DataRow In Dataset.Rows
                        If registro = 1 Then
                            Fecha1.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora1.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion1.Text = CStr(ResumenTabla("ubicacion"))
                            tenor1.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 2 Then
                            Fecha2.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora2.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion2.Text = CStr(ResumenTabla("ubicacion"))
                            tenor2.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 3 Then
                            Fecha3.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora3.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion3.Text = CStr(ResumenTabla("ubicacion"))
                            tenor3.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 4 Then
                            Fecha4.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora4.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion4.Text = CStr(ResumenTabla("ubicacion"))
                            tenor4.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 5 Then
                            Fecha5.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora5.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion5.Text = CStr(ResumenTabla("ubicacion"))
                            tenor5.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 6 Then
                            Fecha6.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora6.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion6.Text = CStr(ResumenTabla("ubicacion"))
                            tenor6.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 7 Then
                            Fecha7.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora7.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion7.Text = CStr(ResumenTabla("ubicacion"))
                            tenor7.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 8 Then
                            Fecha8.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora8.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion8.Text = CStr(ResumenTabla("ubicacion"))
                            tenor8.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 9 Then
                            Fecha9.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora9.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion9.Text = CStr(ResumenTabla("ubicacion"))
                            tenor9.Text = CStr(ResumenTabla("Densidad"))
                        ElseIf registro = 10 Then
                            Fecha10.Text = Convert.ToDateTime(ResumenTabla("UltimaFecha")).ToString("yyyy/MM/dd")
                            Hora10.Text = CStr(ResumenTabla("HoraActualizacion"))
                            Ubicacion10.Text = CStr(ResumenTabla("ubicacion"))
                            tenor10.Text = CStr(ResumenTabla("Densidad"))
                        End If
                        registro = registro + 1
                    Next
                Else
                    LblServer.BackColor = Color.Red
                End If
            Else
                LblServer.Text = "OFF-LINE"
                LblServer.BackColor = Color.Red
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub



    Private Sub limpiarcampos()
        Fecha1.Text = "aaaa/mm/dd"
        Fecha2.Text = "aaaa/mm/dd"
        Fecha3.Text = "aaaa/mm/dd"
        Fecha4.Text = "aaaa/mm/dd"
        Fecha5.Text = "aaaa/mm/dd"
        Fecha6.Text = "aaaa/mm/dd"
        Fecha7.Text = "aaaa/mm/dd"
        Fecha8.Text = "aaaa/mm/dd"
        Fecha9.Text = "aaaa/mm/dd"
        Fecha10.Text = "aaaa/mm/dd"
        Hora1.Text = "00:00"
        Hora2.Text = "00:00"
        Hora3.Text = "00:00"
        Hora4.Text = "00:00"
        Hora5.Text = "00:00"
        Hora6.Text = "00:00"
        Hora7.Text = "00:00"
        Hora8.Text = "00:00"
        Hora9.Text = "00:00"
        Hora10.Text = "00:00"
        Ubicacion1.Text = "Ubicacion"
        Ubicacion2.Text = "Ubicacion"
        Ubicacion3.Text = "Ubicacion"
        Ubicacion4.Text = "Ubicacion"
        Ubicacion5.Text = "Ubicacion"
        Ubicacion6.Text = "Ubicacion"
        Ubicacion7.Text = "Ubicacion"
        Ubicacion8.Text = "Ubicacion"
        Ubicacion9.Text = "Ubicacion"
        Ubicacion10.Text = "Ubicacion"
        tenor1.Text = "0.0"
        tenor2.Text = "0.0"
        tenor3.Text = "0.0"
        tenor4.Text = "0.0"
        tenor5.Text = "0.0"
        tenor6.Text = "0.0"
        tenor7.Text = "0.0"
        tenor8.Text = "0.0"
        tenor9.Text = "0.0"
        tenor10.Text = "0.0"

    End Sub

    Private Sub OcultarCampos()
        Fecha1.Text = ""
        Fecha2.Text = ""
        Fecha3.Text = ""
        Fecha4.Text = ""
        Fecha5.Text = ""
        Fecha6.Text = ""
        Fecha7.Text = ""
        Fecha8.Text = ""
        Fecha9.Text = ""
        Fecha10.Text = ""
        Hora1.Text = ""
        Hora2.Text = ""
        Hora3.Text = ""
        Hora4.Text = ""
        Hora5.Text = ""
        Hora6.Text = ""
        Hora7.Text = ""
        Hora8.Text = ""
        Hora9.Text = ""
        Hora10.Text = ""
        Ubicacion1.Text = ""
        Ubicacion2.Text = ""
        Ubicacion3.Text = ""
        Ubicacion4.Text = ""
        Ubicacion5.Text = ""
        Ubicacion6.Text = ""
        Ubicacion7.Text = ""
        Ubicacion8.Text = ""
        Ubicacion9.Text = ""
        Ubicacion10.Text = ""
        tenor1.Text = ""
        tenor2.Text = ""
        tenor3.Text = ""
        tenor4.Text = ""
        tenor5.Text = ""
        tenor6.Text = ""
        tenor7.Text = ""
        tenor8.Text = ""
        tenor9.Text = ""
        tenor10.Text = ""

    End Sub



    Public Sub TimerOn(ByRef Interval As Short)
        If Interval > 0 Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub

    Public Function TiempoEjecutar(ByVal Tiempo As Integer)
        TiempoEjecutar = ""
        TiempoRestante = Tiempo  ' 1 minutos=60 segundos 
        Timer1.Interval = 1000
        Call TimerOn(1000) ' Hechanos a andar el timer
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If TiempoRestante >= 0 Then
            LblEjecutar.Text = TiempoRestante
            TiempoRestante = TiempoRestante - 1
            lblhora.Text = String.Format("{0:G}", DateTime.Now.ToShortTimeString)
        Else
            If control = 1 Then
                cargarinstantaneas()
            ElseIf control = 2 Then
                cargarDensidades()
                control = 0
            End If
            control = control + 1
            Call TiempoEjecutar(15)
        End If
    End Sub


    Private Sub tenor6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tenor6.Click

    End Sub


    Private Sub lblhora_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblhora.Click

    End Sub

    Private Sub tenor3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tenor3.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class
