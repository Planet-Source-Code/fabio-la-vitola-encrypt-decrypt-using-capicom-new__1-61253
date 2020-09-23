VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView treXML 
      Height          =   2505
      Left            =   240
      TabIndex        =   2
      Top             =   225
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4419
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add NEW node, save and encrypt"
      Height          =   390
      Left            =   4800
      TabIndex        =   1
      Top             =   3195
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   405
      Left            =   5190
      TabIndex        =   0
      Top             =   1200
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XmlDoc As New MSXML2.DOMDocument40
Dim Encryptor As New Codifier

Private Sub Command1_Click()
LoadEncXML
End Sub

Private Sub Command2_Click()
   Dim Nodo As MSXML2.IXMLDOMNode
   Dim NewNode As MSXML2.IXMLDOMNode
   Dim Progr As Integer
   
   With XmlDoc
      Set Nodo = .selectSingleNode("Lavoro/Nodo/Sottonodo")
      Progr = Nodo.childNodes.length + 1
      Set NewNode = CreateNewNode(XmlDoc, Nodo, "Licenza_" & Progr, "")
      NewNode.Text = "Valore" & Progr
      Encryptor.EncryptTextToFile App.Path & "\Prova.xml.enc", XmlDoc.xml, "Prova"
   End With
   LoadEncXML

End Sub

Public Sub CreateNewAttribute(ByVal ParentDoc As MSXML2.DOMDocument40, ByRef ParentNode As MSXML2.IXMLDOMNode, ByVal AttributeNodeName As String, ByVal AttributeValue As String)
   Dim dmyAttrib As MSXML2.IXMLDOMAttribute
   
   Set dmyAttrib = ParentDoc.createAttribute(AttributeNodeName) 'Create New Attribute
   dmyAttrib.Text = AttributeValue                              'Set data to Attribute
   ParentNode.Attributes.setNamedItem dmyAttrib                 'Assign Attribute to parent

End Sub

Public Function CreateNewNode(ByVal ParentDoc As MSXML2.DOMDocument40, ByRef ParentNode As MSXML2.IXMLDOMNode, ByVal NodeName As String, Optional ByVal NodeValue As String) As MSXML2.IXMLDOMNode
   Dim dmyNode As MSXML2.IXMLDOMNode
   'Utilizzo: CreateNewNode Documento, Header, <Nome Nodo>, <Valore Nodo>

   Set dmyNode = ParentDoc.createNode(MSXML2.NODE_ELEMENT, NodeName, vbNullString) 'Create New Node
   dmyNode.Text = NodeValue                                                        'Set data to Node
   ParentNode.appendChild dmyNode                                                  'Assign brand-new node to it's parent
   Set CreateNewNode = dmyNode

End Function

Private Sub Form_Resize()
On Error Resume Next
treXML.Top = 0
treXML.Left = 0
treXML.Height = Me.ScaleHeight


End Sub

Private Sub LoadEncXML()
Dim XMLTestoDec As String
Dim Voce As MSXML2.IXMLDOMNode
Dim N As Integer
Dim Sottonodi As MSXML2.IXMLDOMNode

XMLTestoDec = Encryptor.DecryptFileToText(App.Path & _
   "\Prova.xml.enc", "Prova")                      'Decrypting file into temp string

treXML.Nodes.Clear
With XmlDoc
   .loadXML XMLTestoDec                            'Fill XML object with just decrypted text

   treXML.Nodes.Add , , "Root", .childNodes(1).baseName
   treXML.Nodes.Add "Root", tvwChild, "Nodo", "Nodo"
   treXML.Nodes.Add "Nodo", tvwChild, "SottoNodo", "SottoNodo"
   Set Sottonodi = .selectSingleNode("//Lavoro/Nodo/Sottonodo")
   For N = 0 To Sottonodi.childNodes.length - 1
      Set Voce = Sottonodi.childNodes(N)
      treXML.Nodes.Add "SottoNodo", tvwChild, "N" & N, Voce.baseName
      treXML.Nodes.Add "N" & N, tvwChild, "V" & N, Voce.Text
   Next
   treXML.SetFocus
   treXML.Nodes("N" & N - 1).EnsureVisible
   treXML.Nodes("N" & N - 1).Selected = True
End With


End Sub
