Imports System
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Xml
Imports System.IO
Imports System.Text


Friend Class cXMLLicenseEncryption

    Private ReadOnly cipherText As String
    Private ReadOnly passPhrase As String
    Private ReadOnly saltValue As String
    Private ReadOnly hashAlgorithm As String
    Private ReadOnly passwordIterations As Integer
    Private ReadOnly initVector As String
    Private ReadOnly keySize As Integer

    Private Sub AssymetricEncryptContent(ByVal fileXMLLicensePath As String)
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(fileXMLLicensePath)
        Catch e As Exception
            DLLEventLog.WriteEntry(e.Message, EventLogEntryType.Error)

        End Try
        Dim cspParams As New CspParameters()
        cspParams.KeyContainerName = "s@g@@v1rtu@lc0NTr0lsFTW.c0m"
        ' Create a new RSA key and save it in the container.  This key will encrypt
        ' a symmetric key, which will then be encryped in the XML document.
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            ' Encrypt the "Project" element.
            AssymetricEncrypt(xmlDoc, "Project", "Virtu@lControls", rsaKey, "VirtualControls")
            ' Save the XML document.
            xmlDoc.Save(fileXMLLicensePath)
            '' Display the encrypted XML to the console.
            'MessageBox.Show(xmlDoc.OuterXml)
        Catch e As Exception
            DLLEventLog.WriteEntry(e.Message, EventLogEntryType.Error)
        Finally
            ' Clear the RSA key.
            rsaKey.Clear()
        End Try
    End Sub 'Main

    Private Function AssymetricDecryptContent(ByVal fileXMLLicensePath As String) As XmlDocument
        Dim xmlDoc As New XmlDocument()
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(fileXMLLicensePath)
        Catch e As Exception
            DLLEventLog.WriteEntry(e.Message, EventLogEntryType.Error)
        End Try
        Dim cspParams As New CspParameters()
        cspParams.KeyContainerName = "s@g@@v1rtu@lc0NTr0lsFTW.c0m"
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            AssymetricDecrypt(xmlDoc, rsaKey, "VirtualControls")
            Return xmlDoc
        Catch e As Exception
            Return xmlDoc
        Finally
            rsaKey.Clear()
        End Try
    End Function 'Main

    Private Sub AssymetricEncrypt(ByVal Doc As XmlDocument, ByVal EncryptionElement As String, ByVal EncryptionElementID As String, ByVal Alg As RSA, ByVal KeyName As String)
        ' Check the arguments.
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If EncryptionElement Is Nothing Then
            Throw New ArgumentNullException("EncryptionElement")
        End If
        If EncryptionElementID Is Nothing Then
            Throw New ArgumentNullException("EncryptionElementID")
        End If
        If Alg Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        If KeyName Is Nothing Then
            Throw New ArgumentNullException("KeyName")
        End If
        '//////////////////////////////////////////////
        ' Find the specified element in the XmlDocument
        ' object and create a new XmlElemnt object.
        '//////////////////////////////////////////////
        Dim elementToEncrypt As XmlElement = Doc.GetElementsByTagName(EncryptionElement)(0) '

        ' Throw an XmlException if the element was not found.
        If elementToEncrypt Is Nothing Then
            Throw New XmlException("The specified element was not found")
        End If
        Dim sessionKey As RijndaelManaged = Nothing

        Try
            '////////////////////////////////////////////////
            ' Create a new instance of the EncryptedXml class
            ' and use it to encrypt the XmlElement with the
            ' a new random symmetric key.
            '////////////////////////////////////////////////
            ' Create a 256 bit Rijndael key.
            sessionKey = New RijndaelManaged()
            sessionKey.KeySize = 256

            sessionKey.Key = CreateKey()

            Dim eXml As New EncryptedXml()
            Dim encryptedElement As Byte() = eXml.EncryptData(elementToEncrypt, sessionKey, False)
            '//////////////////////////////////////////////
            ' Construct an EncryptedData object and populate
            ' it with the desired encryption information.
            '//////////////////////////////////////////////
            Dim edElement As New EncryptedData()
            edElement.Type = EncryptedXml.XmlEncElementUrl
            edElement.Id = EncryptionElementID
            ' Create an EncryptionMethod element so that the
            ' receiver knows which algorithm to use for decryption.
            edElement.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncAES256Url)
            ' Encrypt the session key and add it to an EncryptedKey element.
            Dim ek As New EncryptedKey()

            Dim encryptedKey As Byte() = EncryptedXml.EncryptKey(sessionKey.Key, Alg, False)

            ek.CipherData = New CipherData(encryptedKey)

            ek.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncRSA15Url)
            ' Create a new DataReference element
            ' for the KeyInfo element.  This optional
            ' element specifies which EncryptedData
            ' uses this key.  An XML document can have
            ' multiple EncryptedData elements that use
            ' different keys.
            Dim dRef As New DataReference()

            ' Specify the EncryptedData URI.
            dRef.Uri = "#" + EncryptionElementID

            ' Add the DataReference to the EncryptedKey.
            ek.AddReference(dRef)
            ' Add the encrypted key to the
            ' EncryptedData object.
            edElement.KeyInfo.AddClause(New KeyInfoEncryptedKey(ek))
            ' Set the KeyInfo element to specify the
            ' name of the RSA key.
            ' Create a new KeyInfoName element.
            Dim kin As New KeyInfoName()

            ' Specify a name for the key.
            kin.Value = KeyName

            ' Add the KeyInfoName element to the
            ' EncryptedKey object.
            ek.KeyInfo.AddClause(kin)
            ' Add the encrypted element data to the
            ' EncryptedData object.
            edElement.CipherData.CipherValue = encryptedElement
            '//////////////////////////////////////////////////
            ' Replace the element from the original XmlDocument
            ' object with the EncryptedData element.
            '//////////////////////////////////////////////////
            EncryptedXml.ReplaceElement(elementToEncrypt, edElement, False)
            'MessageBox.Show("Docuemnt Encrypted successfully")
        Catch e As Exception
            ' re-throw the exception.
            Throw e
        Finally
            If Not (sessionKey Is Nothing) Then
                sessionKey.Clear()
            End If
        End Try

    End Sub 'Encrypt

    Private Sub AssymetricDecrypt(ByVal Doc As XmlDocument, ByVal Alg As RSA, ByVal KeyName As String)
        ' Check the arguments.  
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If Alg Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        If KeyName Is Nothing Then
            Throw New ArgumentNullException("KeyName")
        End If
        ' Create a new EncryptedXml object.
        Dim exml As New EncryptedXml(Doc)

        ' Add a key-name mapping.
        ' This method can only decrypt documents
        ' that present the specified key name.
        exml.AddKeyNameMapping(KeyName, Alg)

        ' Decrypt the element.
        exml.DecryptDocument()

    End Sub 'Decrypt 

    Private Sub SymetricEncryptContent(ByVal fileXMLLicensePath As String)
        Dim key As RijndaelManaged = Nothing
        Try
            ' Create a new Rijndael key.
            key = New RijndaelManaged()
            key.Key = CreateKey()
            ' Load an XML document.
            Dim xmlDoc As New XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(fileXMLLicensePath)
            ' Encrypt the "creditcard" element.
            SymetricEncrypt(xmlDoc, "Client_Site", key)
            xmlDoc.Save(fileXMLLicensePath)
            'MessageBox.Show("Docuemnt Encrypted successfully")
        Catch e As Exception
            Console.WriteLine(e.Message)
        Finally
            ' Clear the key.
            If Not (key Is Nothing) Then
                key.Clear()
            End If
        End Try

    End Sub

    Private Function SymetricDecryptContent(ByVal fileXMLLicensePath As String) As XmlDocument
        Dim key As RijndaelManaged = Nothing
        Dim xmlDoc As New XmlDocument()
        Try
            ' Create a new Rijndael key.
            key = New RijndaelManaged()
            key.Key = CreateKey()
            ' Load an XML document.
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(fileXMLLicensePath)
            ' Encrypt the "creditcard" element.
            SymetricDecrypt(xmlDoc, key)

            ' We want in runtime to read content and not save it
            'xmlDoc.Save("c:\VirtualControls\test1.xml")
            Return xmlDoc
        Catch e As Exception
            Return xmlDoc
        Finally
            ' Clear the key.
            If Not (key Is Nothing) Then
                key.Clear()
            End If
        End Try
    End Function

    Private Sub SymetricEncrypt(ByVal Doc As XmlDocument, ByVal ElementName As String, ByVal Key As SymmetricAlgorithm)
        ' Check the arguments.  
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If ElementName Is Nothing Then
            Throw New ArgumentNullException("ElementToEncrypt")
        End If
        If Key Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Find the specified element in the XmlDocument
        ' object and create a new XmlElemnt object.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim elementToEncrypt As XmlElement = Doc.GetElementsByTagName(ElementName)(0)

        ' Throw an XmlException if the element was not found.
        If elementToEncrypt Is Nothing Then
            Throw New XmlException("The specified element was not found")
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Create a new instance of the EncryptedXml class 
        ' and use it to encrypt the XmlElement with the 
        ' symmetric key.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim eXml As New EncryptedXml()

        Dim encryptedElement As Byte() = eXml.EncryptData(elementToEncrypt, Key, False)
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Construct an EncryptedData object and populate
        ' it with the desired encryption information.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim edElement As New EncryptedData()
        edElement.Type = EncryptedXml.XmlEncElementUrl
        ' Create an EncryptionMethod element so that the 
        ' receiver knows which algorithm to use for decryption.
        ' Determine what kind of algorithm is being used and
        ' supply the appropriate URL to the EncryptionMethod element.
        Dim encryptionMethod As String = Nothing

        If TypeOf Key Is TripleDES Then
            encryptionMethod = EncryptedXml.XmlEncTripleDESUrl
        ElseIf TypeOf Key Is DES Then
            encryptionMethod = EncryptedXml.XmlEncDESUrl
        End If
        If TypeOf Key Is Rijndael Then
            Select Case Key.KeySize
                Case 128
                    encryptionMethod = EncryptedXml.XmlEncAES128Url
                Case 192
                    encryptionMethod = EncryptedXml.XmlEncAES192Url
                Case 256
                    encryptionMethod = EncryptedXml.XmlEncAES256Url
            End Select
        Else
            ' Throw an exception if the transform is not in the previous categories
            Throw New CryptographicException("The specified algorithm is not supported for XML Encryption.")
        End If

        edElement.EncryptionMethod = New EncryptionMethod(encryptionMethod)
        ' Add the encrypted element data to the 
        ' EncryptedData object.
        edElement.CipherData.CipherValue = encryptedElement
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Replace the element from the original XmlDocument
        ' object with the EncryptedData element.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        EncryptedXml.ReplaceElement(elementToEncrypt, edElement, False)

    End Sub 'Encrypt

    Private Sub SymetricDecrypt(ByVal Doc As XmlDocument, ByVal Alg As SymmetricAlgorithm)
        ' Check the arguments.  
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If Alg Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        ' Find the EncryptedData element in the XmlDocument.
        Dim encryptedElement As XmlElement = Doc.GetElementsByTagName("EncryptedData")(0)

        ' If the EncryptedData element was not found, throw an exception.
        If encryptedElement Is Nothing Then
            Throw New XmlException("The EncryptedData element was not found.")
        End If


        ' Create an EncryptedData object and populate it.
        Dim edElement As New EncryptedData()
        edElement.LoadXml(encryptedElement)
        ' Create a new EncryptedXml object.
        Dim exml As New EncryptedXml()


        ' Decrypt the element using the symmetric key.
        Dim rgbOutput As Byte() = exml.DecryptData(edElement, Alg)
        ' Replace the encryptedData element with the plaintext XML element.
        exml.ReplaceData(encryptedElement, rgbOutput)
    End Sub

    Public Sub New()
        passPhrase = "S@g@FTW_TBH" ' can be any string
        saltValue = "s@1tValue" ' can be any string
        hashAlgorithm = "SHA1" ' can be "MD5"
        passwordIterations = 2  ' can be any number
        initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
        keySize = 256 ' can be 192 or 128
    End Sub

    Private Function CreateKey() As Byte()
        ' Convert strings into byte arrays.
        ' Let us assume that strings only contain ASCII codes.
        ' If strings include Unicode characters, use Unicode, UTF7, or UTF8 
        ' encoding.
        Dim initVectorBytes As Byte()
        initVectorBytes = Encoding.ASCII.GetBytes(initVector)

        Dim saltValueBytes As Byte()
        saltValueBytes = Encoding.ASCII.GetBytes(saltValue)

        ' Convert our plaintext into a byte array.
        ' Let us assume that plaintext contains UTF8-encoded characters.

        Dim password As PasswordDeriveBytes
        password = New PasswordDeriveBytes(passPhrase, _
                                           saltValueBytes, _
                                           hashAlgorithm, _
                                           passwordIterations)

        ' Use the password to generate pseudo-random bytes for the encryption
        ' key. Specify the size of the key in bytes (instead of bits).
        Dim keyBytes As Byte()
        keyBytes = password.GetBytes(keySize / 8)

        ' Create uninitialized Rijndael encryption object.
        CreateKey = keyBytes
    End Function

    Public Sub AssymetricXMLEncryption(ByVal XMLLicensePath As String)
        AssymetricEncryptContent(XMLLicensePath)
    End Sub

    Public Sub SymetricXMLEncryption(ByVal XMLLicensePath As String)
        SymetricEncryptContent(XMLLicensePath)
    End Sub

    Public Function ReadDataFromXML(ByVal TypeOfEncryption As Integer, ByVal XMLLicensePath As String) As String()
        ' 0 for Symmetric Run Time Decryption Of XML
        ' 1 for Asymmetric Run Time Decryption Of XML

        Dim aXmlData As String()
        Dim tempXmlDoc As XmlDocument
        Try
            Select Case TypeOfEncryption
                Case 0
                    tempXmlDoc = SymetricDecryptContent(XMLLicensePath)
                Case 1
                    tempXmlDoc = AssymetricDecryptContent(XMLLicensePath)
                Case Else
                    tempXmlDoc = SymetricDecryptContent(XMLLicensePath)
            End Select

            Dim MacAddressesDecrypted As XmlNodeList = tempXmlDoc.GetElementsByTagName("Mac")
            If MacAddressesDecrypted.Count > 0 Then
                ReDim aXmlData(MacAddressesDecrypted.Count - 1)
                For i = 0 To MacAddressesDecrypted.Count - 1
                    aXmlData(i) = MacStringToMac48String(MacAddressesDecrypted.Item(i).InnerText)
                    ' i = i + 1
                Next
                Return aXmlData
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class 'cEncrytption

