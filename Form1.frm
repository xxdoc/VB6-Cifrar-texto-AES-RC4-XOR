VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cifrar texto AES, RC4 y XOR. Jota"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CIFRAR TEXTO AES, RC4 Y XOR. COMPARACIÓN DE TEXTO CIFRADO. RESULTADOS SE MUESTRAN POR CONSOLA."
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Sub Command1_Click()
Dim x As Integer
'Repito 3 veces
For x = 1 To 3
    If x = 1 Then
        'AES
        AES
    ElseIf x = 2 Then
        'RC4
        RC4
    Else
        'XOR
        cifradoXOR
    End If
Next x

'Muestro el informe
Informe
End Sub

Private Sub Form_Load()
'Setear texto y clave
Clave = "j0t4 !@Y0utu3* 123456"
Texto = "Es un texto de prueba para Yotube."
End Sub
Private Sub AES()
'Encriptar/desencriptar texto usando AES
'Inicia el tiempo
lngStart = GetTickCount()
'Encripto en AES
encriptadoAes = AesEncryptString()
'Desencripto en AES
dencriptadoAes = AesDecryptString(encriptadoAes)
'Obtengo tiempo total de la operacion
lngFinish = GetTickCount()
tiempoAes = lngFinish - lngStart
End Sub
Private Sub RC4()
'Encriptar/desencriptar texto usando RC4
'Inicia el tiempo
lngStart = GetTickCount()
'Encripto en RC4
encriptadoRc4 = ToHexDump(CryptRC4(Texto))
'Desencripto en RC4
dencriptadoRc4 = CryptRC4(FromHexDump(encriptadoRc4))
'Obtengo tiempo total de la operacion
lngFinish = GetTickCount()
tiempoRc4 = lngFinish - lngStart
End Sub
Private Sub cifradoXOR()
'Encriptar/desencriptar texto usando XOR
'Inicia el tiempo
lngStart = GetTickCount()
'Encripto en XOR
encriptadoXor = XOREncryption()
'Desencripto en XOR
dencriptadoXor = XORDecryption(encriptadoXor)
'Obtengo tiempo total de la operacion
lngFinish = GetTickCount()
tiempoXor = lngFinish - lngStart
End Sub
Private Sub Informe()
'Informa los resultados de los cifrados
Debug.Print "******************************"
Debug.Print "Texto a encriptar: " & Texto
Debug.Print "Clave: " & Clave
Debug.Print "******************************"
Debug.Print "Tiempo AES: " & CStr(tiempoAes) & " ms. Tiempo RC4: " & CStr(tiempoRc4) & " ms. Tiempo XOR: " & CStr(tiempoXor) & " ms."
Debug.Print "******************************"
Debug.Print "Texto encriptado AES: " & encriptadoAes
Debug.Print "Texto dencriptado AES: " & dencriptadoAes
Debug.Print "******************************"
Debug.Print "Texto encriptado RC4: " & encriptadoRc4
Debug.Print "Texto dencriptado RC4: " & dencriptadoRc4
Debug.Print "******************************"
Debug.Print "Texto encriptado XOR: " & encriptadoXor
Debug.Print "Texto dencriptado XOR: " & dencriptadoXor
End Sub

