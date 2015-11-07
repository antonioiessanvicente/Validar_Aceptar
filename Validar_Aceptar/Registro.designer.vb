<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegistro
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.gBoxDatosPersonales = New System.Windows.Forms.GroupBox()
        Me.cbNacionalidad = New System.Windows.Forms.ComboBox()
        Me.rbGeneroFemenino = New System.Windows.Forms.RadioButton()
        Me.rbGeneroMasculino = New System.Windows.Forms.RadioButton()
        Me.cbAnyoNacimiento = New System.Windows.Forms.ComboBox()
        Me.cbMesNacimiento = New System.Windows.Forms.ComboBox()
        Me.cbDiaNacimiento = New System.Windows.Forms.ComboBox()
        Me.tbTelefono = New System.Windows.Forms.TextBox()
        Me.tbCorreo = New System.Windows.Forms.TextBox()
        Me.tbApellidos = New System.Windows.Forms.TextBox()
        Me.lblTelefono = New System.Windows.Forms.Label()
        Me.lblNacionalidad = New System.Windows.Forms.Label()
        Me.lblApellidos = New System.Windows.Forms.Label()
        Me.lblFechaNacimiento = New System.Windows.Forms.Label()
        Me.lblGenero = New System.Windows.Forms.Label()
        Me.lblNombre = New System.Windows.Forms.Label()
        Me.lblCorreo = New System.Windows.Forms.Label()
        Me.tbNombre = New System.Windows.Forms.TextBox()
        Me.lblEquipo = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.lblNick = New System.Windows.Forms.Label()
        Me.gBoxDatosUsuario = New System.Windows.Forms.GroupBox()
        Me.tbConfirmarPassword = New System.Windows.Forms.TextBox()
        Me.lblConfirmarPassword = New System.Windows.Forms.Label()
        Me.tbPassword = New System.Windows.Forms.TextBox()
        Me.tbNick = New System.Windows.Forms.TextBox()
        Me.tbEquipo = New System.Windows.Forms.TextBox()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.btnRegistrar = New System.Windows.Forms.Button()
        Me.lblInfoCamposObligatorios = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.gBoxDatosPersonales.SuspendLayout()
        Me.gBoxDatosUsuario.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gBoxDatosPersonales
        '
        Me.gBoxDatosPersonales.Controls.Add(Me.cbNacionalidad)
        Me.gBoxDatosPersonales.Controls.Add(Me.rbGeneroFemenino)
        Me.gBoxDatosPersonales.Controls.Add(Me.rbGeneroMasculino)
        Me.gBoxDatosPersonales.Controls.Add(Me.cbAnyoNacimiento)
        Me.gBoxDatosPersonales.Controls.Add(Me.cbMesNacimiento)
        Me.gBoxDatosPersonales.Controls.Add(Me.cbDiaNacimiento)
        Me.gBoxDatosPersonales.Controls.Add(Me.tbTelefono)
        Me.gBoxDatosPersonales.Controls.Add(Me.tbCorreo)
        Me.gBoxDatosPersonales.Controls.Add(Me.tbApellidos)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblTelefono)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblNacionalidad)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblApellidos)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblFechaNacimiento)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblGenero)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblNombre)
        Me.gBoxDatosPersonales.Controls.Add(Me.lblCorreo)
        Me.gBoxDatosPersonales.Controls.Add(Me.tbNombre)
        Me.gBoxDatosPersonales.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gBoxDatosPersonales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gBoxDatosPersonales.Location = New System.Drawing.Point(12, 12)
        Me.gBoxDatosPersonales.Name = "gBoxDatosPersonales"
        Me.gBoxDatosPersonales.Size = New System.Drawing.Size(325, 224)
        Me.gBoxDatosPersonales.TabIndex = 0
        Me.gBoxDatosPersonales.TabStop = False
        Me.gBoxDatosPersonales.Text = "Datos personales"
        '
        'cbNacionalidad
        '
        Me.cbNacionalidad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbNacionalidad.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cbNacionalidad.FormattingEnabled = True
        Me.cbNacionalidad.Location = New System.Drawing.Point(119, 131)
        Me.cbNacionalidad.Name = "cbNacionalidad"
        Me.cbNacionalidad.Size = New System.Drawing.Size(200, 21)
        Me.cbNacionalidad.TabIndex = 8
        '
        'rbGeneroFemenino
        '
        Me.rbGeneroFemenino.AutoSize = True
        Me.rbGeneroFemenino.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.rbGeneroFemenino.Location = New System.Drawing.Point(198, 108)
        Me.rbGeneroFemenino.Name = "rbGeneroFemenino"
        Me.rbGeneroFemenino.Size = New System.Drawing.Size(71, 17)
        Me.rbGeneroFemenino.TabIndex = 7
        Me.rbGeneroFemenino.TabStop = True
        Me.rbGeneroFemenino.Text = "Femenino"
        Me.rbGeneroFemenino.UseVisualStyleBackColor = True
        '
        'rbGeneroMasculino
        '
        Me.rbGeneroMasculino.AutoSize = True
        Me.rbGeneroMasculino.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.rbGeneroMasculino.Location = New System.Drawing.Point(119, 108)
        Me.rbGeneroMasculino.Name = "rbGeneroMasculino"
        Me.rbGeneroMasculino.Size = New System.Drawing.Size(73, 17)
        Me.rbGeneroMasculino.TabIndex = 6
        Me.rbGeneroMasculino.TabStop = True
        Me.rbGeneroMasculino.Text = "Masculino"
        Me.rbGeneroMasculino.UseVisualStyleBackColor = True
        '
        'cbAnyoNacimiento
        '
        Me.cbAnyoNacimiento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbAnyoNacimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cbAnyoNacimiento.FormattingEnabled = True
        Me.cbAnyoNacimiento.Items.AddRange(New Object() {"1925", "1926", "1927", "1928", "1929", "1930", "1931", "1932", "1933", "1934", "1935", "1936", "1937", "1938", "1939", "1940", "1941", "1942", "1943", "1944", "1945", "1946", "1947", "1948", "1949", "1950", "1951", "1952", "1953", "1954", "1955", "1956", "1957", "1958", "1959", "1960", "1961", "1962", "1963", "1964", "1965", "1966", "1967", "1968", "1969", "1970", "1971", "1972", "1973", "1974", "1975", "1976", "1977", "1978", "1979", "1980", "1981", "1982", "1983", "1984", "1985", "1986", "1987", "1988", "1989", "1990", "1991", "1992", "1993", "1994", "1995", "1996", "1997", "1998", "1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015"})
        Me.cbAnyoNacimiento.Location = New System.Drawing.Point(253, 81)
        Me.cbAnyoNacimiento.Name = "cbAnyoNacimiento"
        Me.cbAnyoNacimiento.Size = New System.Drawing.Size(66, 21)
        Me.cbAnyoNacimiento.TabIndex = 5
        '
        'cbMesNacimiento
        '
        Me.cbMesNacimiento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbMesNacimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cbMesNacimiento.FormattingEnabled = True
        Me.cbMesNacimiento.Items.AddRange(New Object() {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.cbMesNacimiento.Location = New System.Drawing.Point(174, 81)
        Me.cbMesNacimiento.Name = "cbMesNacimiento"
        Me.cbMesNacimiento.Size = New System.Drawing.Size(73, 21)
        Me.cbMesNacimiento.TabIndex = 4
        '
        'cbDiaNacimiento
        '
        Me.cbDiaNacimiento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbDiaNacimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.cbDiaNacimiento.FormattingEnabled = True
        Me.cbDiaNacimiento.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"})
        Me.cbDiaNacimiento.Location = New System.Drawing.Point(119, 81)
        Me.cbDiaNacimiento.Name = "cbDiaNacimiento"
        Me.cbDiaNacimiento.Size = New System.Drawing.Size(49, 21)
        Me.cbDiaNacimiento.TabIndex = 3
        '
        'tbTelefono
        '
        Me.tbTelefono.Location = New System.Drawing.Point(119, 185)
        Me.tbTelefono.Name = "tbTelefono"
        Me.tbTelefono.Size = New System.Drawing.Size(100, 20)
        Me.tbTelefono.TabIndex = 10
        '
        'tbCorreo
        '
        Me.tbCorreo.Location = New System.Drawing.Point(119, 158)
        Me.tbCorreo.Name = "tbCorreo"
        Me.tbCorreo.Size = New System.Drawing.Size(200, 20)
        Me.tbCorreo.TabIndex = 9
        '
        'tbApellidos
        '
        Me.tbApellidos.Location = New System.Drawing.Point(119, 54)
        Me.tbApellidos.Name = "tbApellidos"
        Me.tbApellidos.Size = New System.Drawing.Size(172, 20)
        Me.tbApellidos.TabIndex = 2
        '
        'lblTelefono
        '
        Me.lblTelefono.AutoSize = True
        Me.lblTelefono.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblTelefono.Location = New System.Drawing.Point(6, 190)
        Me.lblTelefono.Name = "lblTelefono"
        Me.lblTelefono.Size = New System.Drawing.Size(49, 13)
        Me.lblTelefono.TabIndex = 10
        Me.lblTelefono.Text = "Teléfono"
        '
        'lblNacionalidad
        '
        Me.lblNacionalidad.AutoSize = True
        Me.lblNacionalidad.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblNacionalidad.Location = New System.Drawing.Point(6, 136)
        Me.lblNacionalidad.Name = "lblNacionalidad"
        Me.lblNacionalidad.Size = New System.Drawing.Size(73, 13)
        Me.lblNacionalidad.TabIndex = 4
        Me.lblNacionalidad.Text = "*Nacionalidad"
        '
        'lblApellidos
        '
        Me.lblApellidos.AutoSize = True
        Me.lblApellidos.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblApellidos.Location = New System.Drawing.Point(6, 59)
        Me.lblApellidos.Name = "lblApellidos"
        Me.lblApellidos.Size = New System.Drawing.Size(49, 13)
        Me.lblApellidos.TabIndex = 3
        Me.lblApellidos.Text = "Apellidos"
        '
        'lblFechaNacimiento
        '
        Me.lblFechaNacimiento.AutoSize = True
        Me.lblFechaNacimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblFechaNacimiento.Location = New System.Drawing.Point(6, 86)
        Me.lblFechaNacimiento.Name = "lblFechaNacimiento"
        Me.lblFechaNacimiento.Size = New System.Drawing.Size(91, 13)
        Me.lblFechaNacimiento.TabIndex = 2
        Me.lblFechaNacimiento.Text = "Fecha nacimiento"
        '
        'lblGenero
        '
        Me.lblGenero.AutoSize = True
        Me.lblGenero.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblGenero.Location = New System.Drawing.Point(6, 110)
        Me.lblGenero.Name = "lblGenero"
        Me.lblGenero.Size = New System.Drawing.Size(42, 13)
        Me.lblGenero.TabIndex = 1
        Me.lblGenero.Text = "Género"
        '
        'lblNombre
        '
        Me.lblNombre.AutoSize = True
        Me.lblNombre.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblNombre.Location = New System.Drawing.Point(6, 32)
        Me.lblNombre.Name = "lblNombre"
        Me.lblNombre.Size = New System.Drawing.Size(48, 13)
        Me.lblNombre.TabIndex = 0
        Me.lblNombre.Text = "*Nombre"
        '
        'lblCorreo
        '
        Me.lblCorreo.AutoSize = True
        Me.lblCorreo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblCorreo.Location = New System.Drawing.Point(6, 163)
        Me.lblCorreo.Name = "lblCorreo"
        Me.lblCorreo.Size = New System.Drawing.Size(97, 13)
        Me.lblCorreo.TabIndex = 5
        Me.lblCorreo.Text = "*Correo electrónico"
        '
        'tbNombre
        '
        Me.tbNombre.Location = New System.Drawing.Point(119, 27)
        Me.tbNombre.Name = "tbNombre"
        Me.tbNombre.Size = New System.Drawing.Size(172, 20)
        Me.tbNombre.TabIndex = 1
        '
        'lblEquipo
        '
        Me.lblEquipo.AutoSize = True
        Me.lblEquipo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblEquipo.Location = New System.Drawing.Point(6, 119)
        Me.lblEquipo.Name = "lblEquipo"
        Me.lblEquipo.Size = New System.Drawing.Size(40, 13)
        Me.lblEquipo.TabIndex = 7
        Me.lblEquipo.Text = "Equipo"
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblPassword.Location = New System.Drawing.Point(6, 67)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(65, 13)
        Me.lblPassword.TabIndex = 8
        Me.lblPassword.Text = "*Contraseña"
        '
        'lblNick
        '
        Me.lblNick.AutoSize = True
        Me.lblNick.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblNick.Location = New System.Drawing.Point(6, 41)
        Me.lblNick.Name = "lblNick"
        Me.lblNick.Size = New System.Drawing.Size(33, 13)
        Me.lblNick.TabIndex = 9
        Me.lblNick.Text = "*Nick"
        '
        'gBoxDatosUsuario
        '
        Me.gBoxDatosUsuario.Controls.Add(Me.tbConfirmarPassword)
        Me.gBoxDatosUsuario.Controls.Add(Me.lblConfirmarPassword)
        Me.gBoxDatosUsuario.Controls.Add(Me.tbPassword)
        Me.gBoxDatosUsuario.Controls.Add(Me.tbNick)
        Me.gBoxDatosUsuario.Controls.Add(Me.lblNick)
        Me.gBoxDatosUsuario.Controls.Add(Me.tbEquipo)
        Me.gBoxDatosUsuario.Controls.Add(Me.lblEquipo)
        Me.gBoxDatosUsuario.Controls.Add(Me.lblPassword)
        Me.gBoxDatosUsuario.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gBoxDatosUsuario.Location = New System.Drawing.Point(12, 242)
        Me.gBoxDatosUsuario.Name = "gBoxDatosUsuario"
        Me.gBoxDatosUsuario.Size = New System.Drawing.Size(325, 159)
        Me.gBoxDatosUsuario.TabIndex = 10
        Me.gBoxDatosUsuario.TabStop = False
        Me.gBoxDatosUsuario.Text = "Datos de Usuario eSports"
        '
        'tbConfirmarPassword
        '
        Me.tbConfirmarPassword.Location = New System.Drawing.Point(119, 90)
        Me.tbConfirmarPassword.Name = "tbConfirmarPassword"
        Me.tbConfirmarPassword.Size = New System.Drawing.Size(200, 20)
        Me.tbConfirmarPassword.TabIndex = 13
        '
        'lblConfirmarPassword
        '
        Me.lblConfirmarPassword.AutoSize = True
        Me.lblConfirmarPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.lblConfirmarPassword.Location = New System.Drawing.Point(6, 93)
        Me.lblConfirmarPassword.Name = "lblConfirmarPassword"
        Me.lblConfirmarPassword.Size = New System.Drawing.Size(111, 13)
        Me.lblConfirmarPassword.TabIndex = 16
        Me.lblConfirmarPassword.Text = "*Confirmar contraseña"
        '
        'tbPassword
        '
        Me.tbPassword.Location = New System.Drawing.Point(119, 64)
        Me.tbPassword.Name = "tbPassword"
        Me.tbPassword.Size = New System.Drawing.Size(200, 20)
        Me.tbPassword.TabIndex = 12
        '
        'tbNick
        '
        Me.tbNick.Location = New System.Drawing.Point(119, 38)
        Me.tbNick.Name = "tbNick"
        Me.tbNick.Size = New System.Drawing.Size(200, 20)
        Me.tbNick.TabIndex = 11
        '
        'tbEquipo
        '
        Me.tbEquipo.Location = New System.Drawing.Point(119, 116)
        Me.tbEquipo.Name = "tbEquipo"
        Me.tbEquipo.Size = New System.Drawing.Size(200, 20)
        Me.tbEquipo.TabIndex = 14
        '
        'btnCancelar
        '
        Me.btnCancelar.Location = New System.Drawing.Point(230, 433)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(107, 23)
        Me.btnCancelar.TabIndex = 25
        Me.btnCancelar.Text = "Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'btnRegistrar
        '
        Me.btnRegistrar.Location = New System.Drawing.Point(107, 433)
        Me.btnRegistrar.Name = "btnRegistrar"
        Me.btnRegistrar.Size = New System.Drawing.Size(107, 23)
        Me.btnRegistrar.TabIndex = 24
        Me.btnRegistrar.Text = "Registrar"
        Me.btnRegistrar.UseVisualStyleBackColor = True
        '
        'lblInfoCamposObligatorios
        '
        Me.lblInfoCamposObligatorios.AutoSize = True
        Me.lblInfoCamposObligatorios.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfoCamposObligatorios.Location = New System.Drawing.Point(13, 417)
        Me.lblInfoCamposObligatorios.Name = "lblInfoCamposObligatorios"
        Me.lblInfoCamposObligatorios.Size = New System.Drawing.Size(259, 13)
        Me.lblInfoCamposObligatorios.TabIndex = 27
        Me.lblInfoCamposObligatorios.Text = "*Los campos marcados con asterisco son obligatorios"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmRegistro
        '
        Me.AcceptButton = Me.btnRegistrar
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 465)
        Me.Controls.Add(Me.lblInfoCamposObligatorios)
        Me.Controls.Add(Me.btnRegistrar)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.gBoxDatosUsuario)
        Me.Controls.Add(Me.gBoxDatosPersonales)
        Me.Name = "frmRegistro"
        Me.Text = "Formulario de registro eSports"
        Me.gBoxDatosPersonales.ResumeLayout(False)
        Me.gBoxDatosPersonales.PerformLayout()
        Me.gBoxDatosUsuario.ResumeLayout(False)
        Me.gBoxDatosUsuario.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents gBoxDatosPersonales As GroupBox
    Friend WithEvents rbGeneroFemenino As RadioButton
    Friend WithEvents rbGeneroMasculino As RadioButton
    Friend WithEvents cbAnyoNacimiento As ComboBox
    Friend WithEvents cbMesNacimiento As ComboBox
    Friend WithEvents cbDiaNacimiento As ComboBox
    Friend WithEvents tbTelefono As TextBox
    Friend WithEvents tbCorreo As TextBox
    Friend WithEvents tbApellidos As TextBox
    Friend WithEvents lblTelefono As Label
    Friend WithEvents lblNacionalidad As Label
    Friend WithEvents lblApellidos As Label
    Friend WithEvents lblFechaNacimiento As Label
    Friend WithEvents lblGenero As Label
    Friend WithEvents lblNombre As Label
    Friend WithEvents lblCorreo As Label
    Friend WithEvents tbNombre As TextBox
    Friend WithEvents lblEquipo As Label
    Friend WithEvents lblPassword As Label
    Friend WithEvents lblNick As Label
    Friend WithEvents gBoxDatosUsuario As GroupBox
    Friend WithEvents tbPassword As TextBox
    Friend WithEvents tbNick As TextBox
    Friend WithEvents tbEquipo As TextBox
    Friend WithEvents btnCancelar As Button
    Friend WithEvents btnRegistrar As Button
    Friend WithEvents tbConfirmarPassword As TextBox
    Friend WithEvents lblConfirmarPassword As Label
    Friend WithEvents cbNacionalidad As ComboBox
    Friend WithEvents lblInfoCamposObligatorios As Label
    Friend WithEvents ErrorProvider1 As ErrorProvider
End Class
