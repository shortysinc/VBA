' @title:  Add-in Enable
' @author: Jorge A. Rivas Córdova
' @e-mail: jorge.shortys@gmail.com

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.

Private Sub changeRegistry()
 ' -----------------------------------------------------------------
 '           *Comprobación previa para que no de error*
 ' -----------------------------------------------------------------
 ' @Descripción: Esto se hace porque si no se comprueba que el
 '               LoadBehaviour está a 3, el Addin realmente no está "cargado"
 '               y no se puede hacer el enable posteriormente
 ' -----------------------------------------------------------------
 If (registryModule.regDoes_Key_Exist(registryModule.HKEY_CURRENT_USER, "Software\Microsoft\Office\15.0\Outlook\Resiliency\DisabledItems")) Then
     
    registryModule.regDelete_A_Key registryModule.HKEY_CURRENT_USER, "Software\Microsoft\Office\15.0\Outlook\Resiliency", "DisabledItems"
    
 Else
    
    ' Se comprueba que está activado en registro este "módulo", ya que si no lo está, Outlook lanza una excepción (Como una excepción en Java no capturada...con esto, lo evitamos).
    ' A esto lo hacemos ya que nuestro objetivo es activar los add-ins a como de lugar y no estar generando más errores.
    ' La comprobación de esto se podría hacer más corta (y no hacer 3 If's) pero así es más fácil cambiar el código.
    
    ' En este if se comprueba si está activado el "Módulo" de DMOutlook2013
    If (registryModule.regDoes_Key_Exist(registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\DMOLAddin")) Then
        Dim iLoadBehavior As Integer
        iLoadBehavior = 0
        iLoadBehavior = registryModule.regQuery_A_Key(registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\DMOLAddin", "LoadBehavior")
                        'If it is not 3, change it to 3
        If (iLoadBehavior <> 3) Then
            registryModule.regCreate_Key_Value registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\DMOLAddin", "LoadBehavior", 3
        End If
    End If
  
    ' En este if se comprueba si está activado el "Módulo" de FileToeDocs2010Addin
    If (registryModule.regDoes_Key_Exist(registryModule.HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin")) Then
        Dim iLoadBehaviorF As Integer
        iLoadBehaviorF = 0
        iLoadBehaviorF = registryModule.regQuery_A_Key(registryModule.HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin", "LoadBehavior")
                        'If it is not 3, change it to 3
        If (iLoadBehaviorF <> 3) Then
            registryModule.regCreate_Key_Value registryModule.HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin", "LoadBehavior", 3
        End If
    End If
   
   ' En este if se comprueba si está activado el "Módulo" de Exchange
    If (registryModule.regDoes_Key_Exist(registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\UmOutlookAddin.FormRegionAddin")) Then
        Dim iLoadBehaviorEx As Integer
        iLoadBehaviorEx = 0
        iLoadBehaviorEx = registryModule.regQuery_A_Key(registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\UmOutlookAddin.FormRegionAddin", "LoadBehavior")
                        'If it is not 3, change it to 3
        If (iLoadBehaviorEx <> 3) Then
            registryModule.regCreate_Key_Value registryModule.HKEY_LOCAL_MACHINE, "Software\Microsoft\Office\Outlook\Addins\UmOutlookAddin.FormRegionAddin", "LoadBehavior", 3
        End If
    End If
    ' HKEY_USERS\.DEFAULT\Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin
    If (registryModule.regDoes_Key_Exist(registryModule.HKEY_USERS, ".DEFAULT\Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin")) Then
        Dim iLoadBehaviorUs As Integer
        iLoadBehaviorUs = 0
        iLoadBehaviorUs = registryModule.regQuery_A_Key(registryModule.HKEY_USERS, ".DEFAULT\Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin", "LoadBehavior")
                        'If it is not 3, change it to 3
        If (iLoadBehaviorUs <> 3) Then
            registryModule.regCreate_Key_Value registryModule.HKEY_USERS, ".DEFAULT\Software\Microsoft\Office\Outlook\Addins\FileToeDocs2010Addin", "LoadBehavior", 3
        End If
    End If
   
 End If
  '  Fin comprobación
End Sub
' Esta función se encarga principalmente de activar los add-ins pero a su vez se encarga de modificar en registro el LoadBehaviour para que no
' salten fallos a la hora de capturar excepciones.
Private Sub EnableAddins()
 ' Variables que necesito para
 Dim MaxAddins As Integer
 Dim Contador As Integer
 Dim DMOutlook As Integer
 ' Número Total de Add-ins
 MaxAddins = Application.COMAddIns.Count
 ' Bucle for que recorre todos los addins y los habilita uno por uno
 For Contador = 1 To MaxAddins
   If (Application.COMAddIns(Contador).Connect = False) And ((Application.COMAddIns(Contador).Description = "DMOutlook2013") Or (Application.COMAddIns(Contador).Description = "File To eDocs DM Outlook 2010 Addin") Or (Application.COMAddIns(Contador).Description = "Microsoft Exchange Add-in")) Then
      Application.COMAddIns(Contador).Connect = True
   End If
     
 Next ' Fin FOR
 
 MsgBox "Add-ins: ENABLED"
 
End Sub
' Main
Sub Start()
    changeRegistry
    EnableAddins
End Sub

