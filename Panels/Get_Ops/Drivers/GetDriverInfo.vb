﻿Imports System.Windows.Forms
Imports System.IO
Imports Microsoft.VisualBasic.ControlChars
Imports Microsoft.Dism
Imports System.Threading
Imports DISMTools.Utilities

Public Class GetDriverInfo

    Dim DriverInfoList As New List(Of DismDriverCollection)
    Public InstalledDriverInfo As DismDriverPackageCollection
    Dim InstalledDriverList As New List(Of DismDriverPackage)
    Dim SearchedDriverList As New List(Of DismDriverPackage)

    Dim CurrentHWTarget As Integer
    Dim CurrentHWFile As Integer = -1        ' This variable gets updated every time an element is selected in the driver packages list box
    Dim JumpTo As Integer = -1               ' This variable gets updated every time a target is specified in the Jump To panel

    Dim ButtonTT As New ToolTip()

    Dim IsInDrvPkgs As Boolean

    Private Sub GetDriverInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        Text = "Get driver information"
                        Label1.Text = Text
                        Label2.Text = "What do you want to get information about?"
                        Label3.Text = "Click here to get information about drivers that you've installed or that came with the Windows image you're servicing"
                        Label4.Text = "Click here to get information about drivers that you want to add to the Windows image you're servicing before proceeding with the driver addition process"
                        Label5.Text = "Ready"
                        Label6.Text = "Add or select a driver package to view its information here"
                        Label7.Text = "Hardware targets"
                        Label8.Text = "Hardware description:"
                        Label10.Text = "Hardware ID:"
                        Label12.Text = "Additional IDs:"
                        Label13.Text = "Compatible IDs:"
                        Label16.Text = "Exclude IDs:"
                        Label17.Text = "Hardware manufacturer:"
                        Label20.Text = "Architecture:"
                        Label21.Text = "Jump to target:"
                        Label22.Text = "Published name:"
                        Label24.Text = "Original file name:"
                        Label26.Text = "Provider name:"
                        Label28.Text = "Is critical to the boot process?"
                        Label30.Text = "Version:"
                        Label31.Text = "Class name:"
                        Label33.Text = "Part of the Windows distribution?"
                        Label36.Text = "Driver information"
                        Label37.Text = "Select an installed driver to view its information here"
                        Label39.Text = "Date:"
                        Label41.Text = "Class description:"
                        Label43.Text = "Class GUID:"
                        Label45.Text = "Driver signature status:"
                        Label47.Text = "Catalog file path:"
                        Label48.Text = "You have configured the background processes to not show all drivers present in this image, which includes drivers part of the Windows distribution, so you may not see the driver you're interested in."
                        Button1.Text = "Add driver..."
                        Button2.Text = "Remove selected"
                        Button3.Text = "Remove all"
                        Button7.Text = "Change"
                        Button8.Text = "Save..."
                        Button9.Text = "View driver file information"
                        LinkLabel1.Text = "<- Go back"
                        InstalledDriverLink.Text = "I want to get information about installed drivers in the image"
                        DriverFileLink.Text = "I want to get information about driver files"
                        ListView1.Columns(0).Text = "Published name"
                        ListView1.Columns(1).Text = "Original file name"
                        OpenFileDialog1.Title = "Locate driver files"
                        SearchBox1.Text = "Type here to search for a driver..."
                    Case "ESN"
                        Text = "Obtener información de controladores"
                        Label1.Text = Text
                        Label2.Text = "¿Acerca de qué le gustaría obtener información?"
                        Label3.Text = "Haga clic aquí para obtener información de controladores que ha instalado o que vengan con la imagen de Windows a la que está dando servicio"
                        Label4.Text = "Haga clic aquí para obtener información de controladores que le gustaría añadir a la imagen de Windows a la que está dando servicio antes de proceder con el proceso de adición de controladores"
                        Label5.Text = "Listo"
                        Label6.Text = "Añada o seleccione un paquete de controlador para ver su información aquí"
                        Label7.Text = "Hardware de destino"
                        Label8.Text = "Descripción de hardware:"
                        Label10.Text = "ID de hardware:"
                        Label12.Text = "Identificadores adicionales:"
                        Label13.Text = "Identificadores compatibles:"
                        Label16.Text = "Identificadores excluidos:"
                        Label17.Text = "Fabricante de hardware:"
                        Label20.Text = "Arquitectura:"
                        Label21.Text = "Saltar a hardware:"
                        Label22.Text = "Nombre publicado:"
                        Label24.Text = "Nombre de archivo original:"
                        Label26.Text = "Nombre de proveedor:"
                        Label28.Text = "¿Es crítico para el proceso de arranque?"
                        Label30.Text = "Versión:"
                        Label31.Text = "Nombre de clase:"
                        Label33.Text = "¿Es parte de la distribución de Windows?"
                        Label36.Text = "Información del controlador"
                        Label37.Text = "Seleccione un controlador instalado para obtener su información aquí"
                        Label39.Text = "Fecha:"
                        Label41.Text = "Descripción de clase:"
                        Label43.Text = "Identificador GUID de clase:"
                        Label45.Text = "Estado de firma del controlador:"
                        Label47.Text = "Ruta del archivo de catálogo:"
                        Label48.Text = "Ha configurado los procesos en segundo plano de manera que no se muestren todos los controladores de esta imagen, que incluye controladores parte de la distribución de Windows, por lo que podría no ver el controlador que le interesa."
                        Button1.Text = "Añadir controlador..."
                        Button2.Text = "Eliminar selección"
                        Button3.Text = "Eliminar todos"
                        Button7.Text = "Cambiar"
                        Button8.Text = "Guardar..."
                        Button9.Text = "Ver información del archivo de controladores"
                        LinkLabel1.Text = "<- Atrás"
                        InstalledDriverLink.Text = "Deseo obtener información acerca de controladores instalados en la imagen"
                        DriverFileLink.Text = "Deseo obtener información acerca de archivos de controladores"
                        ListView1.Columns(0).Text = "Nombre publicado"
                        ListView1.Columns(1).Text = "Nombre de archivo original"
                        OpenFileDialog1.Title = "Ubique los archivos de controladores"
                        SearchBox1.Text = "Escriba aquí para buscar un controlador..."
                    Case "FRA"
                        Text = "Obtenir des informations sur les pilotes"
                        Label1.Text = Text
                        Label2.Text = "Sur quoi souhaitez-vous obtenir des informations ?"
                        Label3.Text = "Cliquez ici pour obtenir des informations sur les pilotes que vous avez installés ou qui sont fournis avec l'image Windows dont vous assurez la maintenance"
                        Label4.Text = "Cliquez ici pour obtenir des informations sur les pilotes que vous souhaitez ajouter à l'image Windows que vous maintenez avant de poursuivre le processus d'ajout de pilote"
                        Label5.Text = "Prêt"
                        Label6.Text = "Ajoutez ou sélectionnez un paquet de pilote pour afficher son information ici"
                        Label7.Text = "Cibles matérielles"
                        Label8.Text = "Description du matériel :"
                        Label10.Text = "ID du matériel :"
                        Label12.Text = "ID supplémentaires :"
                        Label13.Text = "ID compatibles :"
                        Label16.Text = "ID d'exclusion :"
                        Label17.Text = "Fabricant de matériel :"
                        Label20.Text = "Architecture :"
                        Label21.Text = "Sauter à la cible :"
                        Label22.Text = "Nom publié :"
                        Label24.Text = "Nom du fichier original :"
                        Label26.Text = "Nom du prestataire :"
                        Label28.Text = "Est-il essentiel au processus de démarrage ?"
                        Label30.Text = "Version :"
                        Label31.Text = "Nom de classe :"
                        Label33.Text = "Fait-il partie de la distribution Windows ?"
                        Label36.Text = "Information sur le pilote"
                        Label37.Text = "Sélectionnez un pilote installé pour afficher ses informations ici"
                        Label39.Text = "Date :"
                        Label41.Text = "Description de classe :"
                        Label43.Text = "GUID de classe :"
                        Label45.Text = "État de la signature du pilote :"
                        Label47.Text = "Chemin d'accès au fichier de catalogue :"
                        Label48.Text = "Vous avez configuré les processus en arrière plan de manière à ne pas afficher tous les pilotes présents dans cette image, ce qui inclut les pilotes faisant partie de la distribution Windows. Il est donc possible que vous ne voyiez pas le pilote qui vous intéresse."
                        Button1.Text = "Ajouter un pilote..."
                        Button2.Text = "Supprimer la sélection"
                        Button3.Text = "Supprimer tout"
                        Button7.Text = "Changer"
                        Button8.Text = "Sauvegarder..."
                        Button9.Text = "Voir les informations sur le fichier pilote"
                        LinkLabel1.Text = "<- Retourner"
                        InstalledDriverLink.Text = "Je souhaite obtenir des informations sur les pilotes installés dans l'image."
                        DriverFileLink.Text = "Je souhaite obtenir des informations sur les fichiers pilotes"
                        ListView1.Columns(0).Text = "Nom publié"
                        ListView1.Columns(1).Text = "Nom du fichier original"
                        OpenFileDialog1.Title = "Localiser les fichiers pilotes"
                        SearchBox1.Text = "Tapez ici pour rechercher un pilote..."
                    Case "PTB", "PTG"
                        Text = "Obter informações do controlador"
                        Label1.Text = Text
                        Label2.Text = "Sobre o que é que pretende obter informações?"
                        Label3.Text = "Clique aqui para obter informações sobre os controladores que instalou ou que vieram com a imagem do Windows que está a reparar"
                        Label4.Text = "Clique aqui para obter informações sobre os controladores que pretende adicionar à imagem do Windows que está a reparar antes de prosseguir com o processo de adição de controladores"
                        Label5.Text = "Pronto"
                        Label6.Text = "Adicione ou seleccione um pacote de controladores para ver as suas informações aqui"
                        Label7.Text = "Alvos de hardware"
                        Label8.Text = "Descrição do hardware:"
                        Label10.Text = "ID do hardware:"
                        Label12.Text = "IDs adicionais:"
                        Label13.Text = "IDs compatíveis:"
                        Label16.Text = "Excluir IDs:"
                        Label17.Text = "Fabricante do hardware:"
                        Label20.Text = "Arquitetura:"
                        Label21.Text = "Saltar para o alvo:"
                        Label22.Text = "Nome publicado:"
                        Label24.Text = "Nome do ficheiro original:"
                        Label26.Text = "Nome do fornecedor:"
                        Label28.Text = "É crítico para o processo de arranque?"
                        Label30.Text = "Versão:"
                        Label31.Text = "Nome da classe:"
                        Label33.Text = "Parte da distribuição do Windows?"
                        Label36.Text = "Informações do controlador"
                        Label37.Text = "Seleccione um controlador instalado para ver as suas informações aqui"
                        Label39.Text = "Data:"
                        Label41.Text = "Descrição da classe:"
                        Label43.Text = "GUID da classe:"
                        Label45.Text = "Estado da assinatura do controlador:"
                        Label47.Text = "Caminho do ficheiro de catálogo:"
                        Label48.Text = "Configurou os processos em segundo plano para não mostrar todos os controladores presentes nesta imagem, o que inclui controladores que fazem parte da distribuição do Windows, pelo que poderá não ver o controlador em que está interessado."
                        Button1.Text = "Adicionar controlador..."
                        Button2.Text = "Remover selecionado"
                        Button3.Text = "Remover todos"
                        Button7.Text = "Alterar"
                        Button8.Text = "Guardar..."
                        Button9.Text = "Ver informações do ficheiro do controlador"
                        LinkLabel1.Text = "<- Voltar atrás"
                        InstalledDriverLink.Text = "Quero obter informações sobre os controladores instalados na imagem"
                        DriverFileLink.Text = "Pretendo obter informações sobre ficheiros de controladores"
                        ListView1.Columns(0).Text = "Nome publicado"
                        SearchBox1.Text = "Digite aqui para pesquisar um controlador..."
                    Case "ITA"
                        Text = "Ottieni informazioni sul conducente"
                        Label1.Text = Text
                        Label2.Text = "Su cosa si desidera ottenere informazioni?"
                        Label3.Text = "Fare clic qui per ottenere informazioni sui driver installati o forniti con l'immagine di Windows che si sta revisionando"
                        Label4.Text = "Fare clic qui per ottenere informazioni sui driver che si desidera aggiungere all'immagine di Windows in assistenza prima di procedere con il processo di aggiunta dei driver"
                        Label5.Text = "Pronto"
                        Label6.Text = "Aggiungere o selezionare un pacchetto di driver per visualizzarne le informazioni"
                        Label7.Text = "Obiettivi hardware"
                        Label8.Text = "Descrizione hardware:"
                        Label10.Text = "ID hardware:"
                        Label12.Text = "ID aggiuntivi:"
                        Label13.Text = "ID compatibili:"
                        Label16.Text = "Escludere ID:"
                        Label17.Text = "Produttore hardware:"
                        Label20.Text = "Architettura:"
                        Label21.Text = "Salto all'obiettivo:"
                        Label22.Text = "Nome pubblicato:"
                        Label24.Text = "Nome file originale:"
                        Label26.Text = "Nome del fornitore:"
                        Label28.Text = "È fondamentale per il processo di avvio?"
                        Label30.Text = "Versione:"
                        Label31.Text = "Nome della classe:"
                        Label33.Text = "Parte della distribuzione di Windows?"
                        Label36.Text = "Informazioni sul driver"
                        Label37.Text = "Selezionare un driver installato per visualizzarne le informazioni"
                        Label39.Text = "Data:"
                        Label41.Text = "Descrizione della classe:"
                        Label43.Text = "GUID della classe:"
                        Label45.Text = "Stato della firma del driver:"
                        Label47.Text = "Percorso del file di catalogo:"
                        Label48.Text = "I processi in background sono stati configurati in modo da non mostrare tutti i driver presenti in questa immagine, che include i driver che fanno parte della distribuzione di Windows, quindi è possibile che non venga visualizzato il driver di interesse."
                        Button1.Text = "Aggiungi driver..."
                        Button2.Text = "Rimuovi selezionati"
                        Button3.Text = "Rimuovi tutti"
                        Button7.Text = "Cambia"
                        Button8.Text = "Salva..."
                        Button9.Text = "Visualizza informazioni sul file del driver"
                        LinkLabel1.Text = "<- Indietro"
                        InstalledDriverLink.Text = "Voglio ottenere informazioni sui driver installati nell'immagine"
                        DriverFileLink.Text = "Desidero ottenere informazioni sui file dei driver"
                        ListView1.Columns(0).Text = "Nome del file pubblicato"
                        ListView1.Columns(1).Text = "Nome file originale"
                        OpenFileDialog1.Title = "Individuazione dei file di driver"
                        SearchBox1.Text = "Digitare qui per cercare un driver..."
                End Select
            Case 1
                Text = "Get driver information"
                Label1.Text = Text
                Label2.Text = "What do you want to get information about?"
                Label3.Text = "Click here to get information about drivers that you've installed or that came with the Windows image you're servicing"
                Label4.Text = "Click here to get information about drivers that you want to add to the Windows image you're servicing before proceeding with the driver addition process"
                Label5.Text = "Ready"
                Label6.Text = "Add or select a driver package to view its information here"
                Label7.Text = "Hardware targets"
                Label8.Text = "Hardware description:"
                Label10.Text = "Hardware ID:"
                Label12.Text = "Additional IDs:"
                Label13.Text = "Compatible IDs:"
                Label16.Text = "Exclude IDs:"
                Label17.Text = "Hardware manufacturer:"
                Label20.Text = "Architecture:"
                Label21.Text = "Jump to target:"
                Label22.Text = "Published name:"
                Label24.Text = "Original file name:"
                Label26.Text = "Provider name:"
                Label28.Text = "Is critical to the boot process?"
                Label30.Text = "Version:"
                Label31.Text = "Class name:"
                Label33.Text = "Part of the Windows distribution?"
                Label36.Text = "Driver information"
                Label37.Text = "Select an installed driver to view its information here"
                Label39.Text = "Date:"
                Label41.Text = "Class description:"
                Label43.Text = "Class GUID:"
                Label45.Text = "Driver signature status:"
                Label47.Text = "Catalog file path:"
                Label48.Text = "You have configured the background processes to not show all drivers present in this image, which includes drivers part of the Windows distribution, so you may not see the driver you're interested in."
                Button1.Text = "Add driver..."
                Button2.Text = "Remove selected"
                Button3.Text = "Remove all"
                Button7.Text = "Change"
                Button8.Text = "Save..."
                Button9.Text = "View driver file information"
                LinkLabel1.Text = "<- Go back"
                InstalledDriverLink.Text = "I want to get information about installed drivers in the image"
                DriverFileLink.Text = "I want to get information about driver files"
                ListView1.Columns(0).Text = "Published name"
                ListView1.Columns(1).Text = "Original file name"
                OpenFileDialog1.Title = "Locate driver files"
                SearchBox1.Text = "Type here to search for a driver..."
            Case 2
                Text = "Obtener información de controladores"
                Label1.Text = Text
                Label2.Text = "¿Acerca de qué le gustaría obtener información?"
                Label3.Text = "Haga clic aquí para obtener información de controladores que ha instalado o que vengan con la imagen de Windows a la que está dando servicio"
                Label4.Text = "Haga clic aquí para obtener información de controladores que le gustaría añadir a la imagen de Windows a la que está dando servicio antes de proceder con el proceso de adición de controladores"
                Label5.Text = "Listo"
                Label6.Text = "Añada o seleccione un paquete de controlador para ver su información aquí"
                Label7.Text = "Hardware de destino"
                Label8.Text = "Descripción de hardware:"
                Label10.Text = "ID de hardware:"
                Label12.Text = "Identificadores adicionales:"
                Label13.Text = "Identificadores compatibles:"
                Label16.Text = "Identificadores excluidos:"
                Label17.Text = "Fabricante de hardware:"
                Label20.Text = "Arquitectura:"
                Label21.Text = "Saltar a hardware:"
                Label22.Text = "Nombre publicado:"
                Label24.Text = "Nombre de archivo original:"
                Label26.Text = "Nombre de proveedor:"
                Label28.Text = "¿Es crítico para el proceso de arranque?"
                Label30.Text = "Versión:"
                Label31.Text = "Nombre de clase:"
                Label33.Text = "¿Es parte de la distribución de Windows?"
                Label36.Text = "Información del controlador"
                Label37.Text = "Seleccione un controlador instalado para obtener su información aquí"
                Label39.Text = "Fecha:"
                Label41.Text = "Descripción de clase:"
                Label43.Text = "Identificador GUID de clase:"
                Label45.Text = "Estado de firma del controlador:"
                Label47.Text = "Ruta del archivo de catálogo:"
                Label48.Text = "Ha configurado los procesos en segundo plano de manera que no se muestren todos los controladores de esta imagen, que incluye controladores parte de la distribución de Windows, por lo que podría no ver el controlador que le interesa."
                Button1.Text = "Añadir controlador..."
                Button2.Text = "Eliminar selección"
                Button3.Text = "Eliminar todos"
                Button7.Text = "Cambiar"
                Button8.Text = "Guardar..."
                Button9.Text = "Ver información del archivo de controladores"
                LinkLabel1.Text = "<- Atrás"
                InstalledDriverLink.Text = "Deseo obtener información acerca de controladores instalados en la imagen"
                DriverFileLink.Text = "Deseo obtener información acerca de archivos de controladores"
                ListView1.Columns(0).Text = "Nombre publicado"
                ListView1.Columns(1).Text = "Nombre de archivo original"
                OpenFileDialog1.Title = "Ubique los archivos de controladores"
                SearchBox1.Text = "Escriba aquí para buscar un controlador..."
            Case 3
                Text = "Obtenir des informations sur les pilotes"
                Label1.Text = Text
                Label2.Text = "Sur quoi souhaitez-vous obtenir des informations ?"
                Label3.Text = "Cliquez ici pour obtenir des informations sur les pilotes que vous avez installés ou qui sont fournis avec l'image Windows dont vous assurez la maintenance"
                Label4.Text = "Cliquez ici pour obtenir des informations sur les pilotes que vous souhaitez ajouter à l'image Windows que vous maintenez avant de poursuivre le processus d'ajout de pilote"
                Label5.Text = "Prêt"
                Label6.Text = "Ajoutez ou sélectionnez un paquet de pilote pour afficher son information ici"
                Label7.Text = "Cibles matérielles"
                Label8.Text = "Description du matériel :"
                Label10.Text = "ID du matériel :"
                Label12.Text = "ID supplémentaires :"
                Label13.Text = "ID compatibles :"
                Label16.Text = "ID d'exclusion :"
                Label17.Text = "Fabricant de matériel :"
                Label20.Text = "Architecture :"
                Label21.Text = "Sauter à la cible :"
                Label22.Text = "Nom publié :"
                Label24.Text = "Nom du fichier original :"
                Label26.Text = "Nom du prestataire :"
                Label28.Text = "Est-il essentiel au processus de démarrage ?"
                Label30.Text = "Version :"
                Label31.Text = "Nom de classe :"
                Label33.Text = "Fait-il partie de la distribution Windows ?"
                Label36.Text = "Information sur le pilote"
                Label37.Text = "Sélectionnez un pilote installé pour afficher ses informations ici"
                Label39.Text = "Date :"
                Label41.Text = "Description de classe :"
                Label43.Text = "GUID de classe :"
                Label45.Text = "État de la signature du pilote :"
                Label47.Text = "Chemin d'accès au fichier de catalogue :"
                Label48.Text = "Vous avez configuré les processus en arrière plan de manière à ne pas afficher tous les pilotes présents dans cette image, ce qui inclut les pilotes faisant partie de la distribution Windows. Il est donc possible que vous ne voyiez pas le pilote qui vous intéresse."
                Button1.Text = "Ajouter un pilote..."
                Button2.Text = "Supprimer la sélection"
                Button3.Text = "Supprimer tout"
                Button7.Text = "Changer"
                Button8.Text = "Sauvegarder..."
                Button9.Text = "Voir les informations sur le fichier pilote"
                LinkLabel1.Text = "<- Retourner"
                InstalledDriverLink.Text = "Je souhaite obtenir des informations sur les pilotes installés dans l'image."
                DriverFileLink.Text = "Je souhaite obtenir des informations sur les fichiers pilotes"
                ListView1.Columns(0).Text = "Nom publié"
                ListView1.Columns(1).Text = "Nom du fichier original"
                OpenFileDialog1.Title = "Localiser les fichiers pilotes"
                SearchBox1.Text = "Tapez ici pour rechercher un pilote..."
            Case 4
                Text = "Obter informações do controlador"
                Label1.Text = Text
                Label2.Text = "Sobre o que é que pretende obter informações?"
                Label3.Text = "Clique aqui para obter informações sobre os controladores que instalou ou que vieram com a imagem do Windows que está a reparar"
                Label4.Text = "Clique aqui para obter informações sobre os controladores que pretende adicionar à imagem do Windows que está a reparar antes de prosseguir com o processo de adição de controladores"
                Label5.Text = "Pronto"
                Label6.Text = "Adicione ou seleccione um pacote de controladores para ver as suas informações aqui"
                Label7.Text = "Alvos de hardware"
                Label8.Text = "Descrição do hardware:"
                Label10.Text = "ID do hardware:"
                Label12.Text = "IDs adicionais:"
                Label13.Text = "IDs compatíveis:"
                Label16.Text = "Excluir IDs:"
                Label17.Text = "Fabricante do hardware:"
                Label20.Text = "Arquitetura:"
                Label21.Text = "Saltar para o alvo:"
                Label22.Text = "Nome publicado:"
                Label24.Text = "Nome do ficheiro original:"
                Label26.Text = "Nome do fornecedor:"
                Label28.Text = "É crítico para o processo de arranque?"
                Label30.Text = "Versão:"
                Label31.Text = "Nome da classe:"
                Label33.Text = "Parte da distribuição do Windows?"
                Label36.Text = "Informações do controlador"
                Label37.Text = "Seleccione um controlador instalado para ver as suas informações aqui"
                Label39.Text = "Data:"
                Label41.Text = "Descrição da classe:"
                Label43.Text = "GUID da classe:"
                Label45.Text = "Estado da assinatura do controlador:"
                Label47.Text = "Caminho do ficheiro de catálogo:"
                Label48.Text = "Configurou os processos em segundo plano para não mostrar todos os controladores presentes nesta imagem, o que inclui controladores que fazem parte da distribuição do Windows, pelo que poderá não ver o controlador em que está interessado."
                Button1.Text = "Adicionar controlador..."
                Button2.Text = "Remover selecionado"
                Button3.Text = "Remover todos"
                Button7.Text = "Alterar"
                Button8.Text = "Guardar..."
                Button9.Text = "Ver informações do ficheiro do controlador"
                LinkLabel1.Text = "<- Voltar atrás"
                InstalledDriverLink.Text = "Quero obter informações sobre os controladores instalados na imagem"
                DriverFileLink.Text = "Pretendo obter informações sobre ficheiros de controladores"
                ListView1.Columns(0).Text = "Nome publicado"
                SearchBox1.Text = "Digite aqui para pesquisar um controlador..."
            Case 5
                Text = "Ottieni informazioni sul conducente"
                Label1.Text = Text
                Label2.Text = "Su cosa si desidera ottenere informazioni?"
                Label3.Text = "Fare clic qui per ottenere informazioni sui driver installati o forniti con l'immagine di Windows che si sta revisionando"
                Label4.Text = "Fare clic qui per ottenere informazioni sui driver che si desidera aggiungere all'immagine di Windows in assistenza prima di procedere con il processo di aggiunta dei driver"
                Label5.Text = "Pronto"
                Label6.Text = "Aggiungere o selezionare un pacchetto di driver per visualizzarne le informazioni"
                Label7.Text = "Obiettivi hardware"
                Label8.Text = "Descrizione hardware:"
                Label10.Text = "ID hardware:"
                Label12.Text = "ID aggiuntivi:"
                Label13.Text = "ID compatibili:"
                Label16.Text = "Escludere ID:"
                Label17.Text = "Produttore hardware:"
                Label20.Text = "Architettura:"
                Label21.Text = "Salto all'obiettivo:"
                Label22.Text = "Nome pubblicato:"
                Label24.Text = "Nome file originale:"
                Label26.Text = "Nome del fornitore:"
                Label28.Text = "È fondamentale per il processo di avvio?"
                Label30.Text = "Versione:"
                Label31.Text = "Nome della classe:"
                Label33.Text = "Parte della distribuzione di Windows?"
                Label36.Text = "Informazioni sul driver"
                Label37.Text = "Selezionare un driver installato per visualizzarne le informazioni"
                Label39.Text = "Data:"
                Label41.Text = "Descrizione della classe:"
                Label43.Text = "GUID della classe:"
                Label45.Text = "Stato della firma del driver:"
                Label47.Text = "Percorso del file di catalogo:"
                Label48.Text = "I processi in background sono stati configurati in modo da non mostrare tutti i driver presenti in questa immagine, che include i driver che fanno parte della distribuzione di Windows, quindi è possibile che non venga visualizzato il driver di interesse."
                Button1.Text = "Aggiungi driver..."
                Button2.Text = "Rimuovi selezionati"
                Button3.Text = "Rimuovi tutti"
                Button7.Text = "Cambia"
                Button8.Text = "Salva..."
                Button9.Text = "Visualizza informazioni sul file del driver"
                LinkLabel1.Text = "<- Indietro"
                InstalledDriverLink.Text = "Voglio ottenere informazioni sui driver installati nell'immagine"
                DriverFileLink.Text = "Desidero ottenere informazioni sui file dei driver"
                ListView1.Columns(0).Text = "Nome del file pubblicato"
                ListView1.Columns(1).Text = "Nome file originale"
                OpenFileDialog1.Title = "Individuazione dei file di driver"
                SearchBox1.Text = "Digitare qui per cercare un driver..."
        End Select
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Win10Title.BackColor = Color.FromArgb(48, 48, 48)
            BackColor = Color.FromArgb(31, 31, 31)
            ForeColor = Color.White
            ListBox1.BackColor = Color.FromArgb(31, 31, 31)
            ListView1.BackColor = Color.FromArgb(31, 31, 31)
            ComboBox1.BackColor = Color.FromArgb(31, 31, 31)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Win10Title.BackColor = Color.White
            BackColor = Color.FromArgb(238, 238, 242)
            ForeColor = Color.Black
            ListBox1.BackColor = Color.FromArgb(238, 238, 242)
            ListView1.BackColor = Color.FromArgb(238, 238, 242)
            ComboBox1.BackColor = Color.FromArgb(238, 238, 242)
        End If
        ListBox1.ForeColor = ForeColor
        ListView1.ForeColor = ForeColor
        ComboBox1.ForeColor = ForeColor
        SearchBox1.BackColor = BackColor
        SearchBox1.ForeColor = ForeColor
        SearchPic.Image = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), My.Resources.search_dark, My.Resources.search_light)
        If Environment.OSVersion.Version.Major = 10 Then
            Text = ""
            Win10Title.Visible = True
        End If
        Dim handle As IntPtr = MainForm.GetWindowHandle(Me)
        If MainForm.IsWindowsVersionOrGreater(10, 0, 18362) Then MainForm.EnableDarkTitleBar(handle, MainForm.BackColor = Color.FromArgb(48, 48, 48))
        InstalledDriverList.Clear()
        SearchedDriverList.Clear()
        ListView1.Items.Clear()
        If InstalledDriverInfo.Count > 0 Then
            For Each DriverPackage As DismDriverPackage In InstalledDriverInfo
                InstalledDriverList.Add(DriverPackage)
                ListView1.Items.Add(New ListViewItem(New String() {DriverPackage.PublishedName, Path.GetFileName(DriverPackage.OriginalFileName)}))
            Next
        End If

        ' Detect if the "Detect all drivers" option is checked and act accordingly
        Panel6.Visible = MainForm.AllDrivers = False

        ' Forcefully hide that panel if the driver packages panel is visible
        If IsInDrvPkgs Then Panel6.Visible = False

        Button9.Visible = IsInDrvPkgs
        Button9.Enabled = False

        ' Switch to the selection panels
        Panel4.Visible = False
        Panel7.Visible = True
        DrvPackageInfoPanel.Visible = False
        NoDrvPanel.Visible = True

        SearchBox1.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        ListBox1.Items.Add(OpenFileDialog1.FileName)
        Button3.Enabled = True
        Button8.Enabled = True
        GetDriverInformation()
    End Sub

    Private Sub InstalledDriverLink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles InstalledDriverLink.LinkClicked
        MenuPanel.Visible = False
        DriverInfoPanel.Visible = True
        InfoFromInstalledDrvsPanel.Visible = True
        InfoFromDrvPackagesPanel.Visible = False

        ' Detect if the "Detect all drivers" option is checked and act accordingly
        Panel6.Visible = MainForm.AllDrivers = False

        Label5.Visible = False
        IsInDrvPkgs = False
        Button8.Enabled = True
        Button9.Visible = False
    End Sub

    Private Sub DriverFileLink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles DriverFileLink.LinkClicked
        MenuPanel.Visible = False
        DriverInfoPanel.Visible = True
        InfoFromInstalledDrvsPanel.Visible = False
        InfoFromDrvPackagesPanel.Visible = True
        Panel6.Visible = False
        Label5.Visible = True
        IsInDrvPkgs = True
        Button8.Enabled = ListBox1.Items.Count > 0
        Button9.Visible = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MenuPanel.Visible = True
        DriverInfoPanel.Visible = False
    End Sub

    Sub GetDriverInformation()
        DriverInfoList.Clear()
        Try
            ' Background processes need to have completed before showing information
            If MainForm.ImgBW.IsBusy Then
                Dim msg As String = ""
                Select Case MainForm.Language
                    Case 0
                        Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                            Case "ENU", "ENG"
                                msg = "Background processes need to have completed before showing package information. We'll wait until they have completed"
                            Case "ESN"
                                msg = "Los procesos en segundo plano deben haber completado antes de obtener información del paquete. Esperaremos hasta que hayan completado"
                            Case "FRA"
                                msg = "Les processus en plan doivent être terminés avant d'afficher les paquets. Nous attendrons qu'ils soient terminés"
                            Case "PTB", "PTG"
                                msg = "Os processos em segundo plano precisam de ser concluídos antes de mostrar as informações dos pacotes. Esperamos até que estejam concluídos"
                            Case "ITA"
                                msg = "I processi in secondo piano devono essere completati prima di mostrare le informazioni sul pacchetto. Aspetteremo che siano completati"
                        End Select
                    Case 1
                        msg = "Background processes need to have completed before showing package information. We'll wait until they have completed"
                    Case 2
                        msg = "Los procesos en segundo plano deben haber completado antes de obtener información del paquete. Esperaremos hasta que hayan completado"
                    Case 3
                        msg = "Les processus en plan doivent être terminés avant d'afficher les paquets. Nous attendrons qu'ils soient terminés"
                    Case 4
                        msg = "Os processos em segundo plano precisam de ser concluídos antes de mostrar as informações dos pacotes. Esperamos até que estejam concluídos"
                    Case 5
                        msg = "I processi in secondo piano devono essere completati prima di mostrare le informazioni sul pacchetto. Aspetteremo che siano completati"
                End Select
                MsgBox(msg, vbOKOnly + vbInformation, Label1.Text)
                Select Case MainForm.Language
                    Case 0
                        Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                            Case "ENU", "ENG"
                                Label5.Text = "Waiting for background processes to finish..."
                            Case "ESN"
                                Label5.Text = "Esperando a que terminen los procesos en segundo plano..."
                            Case "FRA"
                                Label5.Text = "Attente de la fin des processus en arrière plan..."
                            Case "PTB", "PTG"
                                Label5.Text = "À espera que os processos em segundo plano terminem..."
                            Case "ITA"
                                Label5.Text = "In attesa che i processi in secondo piano finiscano..."
                        End Select
                    Case 1
                        Label5.Text = "Waiting for background processes to finish..."
                    Case 2
                        Label5.Text = "Esperando a que terminen los procesos en segundo plano..."
                    Case 3
                        Label5.Text = "Attente de la fin des processus en arrière plan..."
                    Case 4
                        Label5.Text = "À espera que os processos em segundo plano terminem..."
                    Case 5
                        Label5.Text = "In attesa che i processi in secondo piano finiscano..."
                End Select
                While MainForm.ImgBW.IsBusy
                    Application.DoEvents()
                    Thread.Sleep(500)
                End While
            End If
            If MainForm.MountedImageDetectorBW.IsBusy Then
                MainForm.MountedImageDetectorBWRestarterTimer.Enabled = False
                MainForm.MountedImageDetectorBW.CancelAsync()
                While MainForm.MountedImageDetectorBW.IsBusy
                    Application.DoEvents()
                    Thread.Sleep(500)
                End While
            End If
            MainForm.WatcherTimer.Enabled = False
            If MainForm.WatcherBW.IsBusy Then MainForm.WatcherBW.CancelAsync()
            While MainForm.WatcherBW.IsBusy
                Application.DoEvents()
                Thread.Sleep(100)
            End While
            Select Case MainForm.Language
                Case 0
                    Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                        Case "ENU", "ENG"
                            Label5.Text = "Preparing driver information processes..."
                        Case "ESN"
                            Label5.Text = "Preparando procesos de información de controladores..."
                        Case "FRA"
                            Label5.Text = "Préparation des processus d'information des pilotes en cours..."
                        Case "PTB", "PTG"
                            Label5.Text = "Preparar os processos de informação dos controladores..."
                        Case "ITA"
                            Label5.Text = "Preparazione per ottenere le informazioni del driver..."
                    End Select
                Case 1
                    Label5.Text = "Preparing driver information processes..."
                Case 2
                    Label5.Text = "Preparando procesos de información de controladores..."
                Case 3
                    Label5.Text = "Préparation des processus d'information des pilotes en cours..."
                Case 4
                    Label5.Text = "Preparar os processos de informação dos controladores..."
                Case 5
                    Label5.Text = "Preparazione per ottenere le informazioni del driver..."
            End Select
            Application.DoEvents()
            DismApi.Initialize(DismLogLevel.LogErrors)
            Using imgSession As DismSession = If(MainForm.OnlineManagement, DismApi.OpenOnlineSession(), DismApi.OpenOfflineSession(MainForm.MountDir))
                For Each drvFile In ListBox1.Items
                    If File.Exists(drvFile) Then
                        Select Case MainForm.Language
                            Case 0
                                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                                    Case "ENU", "ENG"
                                        Label5.Text = "Getting information from driver file " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "This may take some time and the program may temporarily freeze"
                                    Case "ESN"
                                        Label5.Text = "Obteniendo información del archivo de controlador " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Esto puede llevar algo de tiempo y el programa podría congelarse temporalmente"
                                    Case "FRA"
                                        Label5.Text = "Obtention des informations du fichier pilote " & Quote & Path.GetFileName(drvFile) & Quote & " en cours..." & CrLf & "Cette opération peut prendre un certain temps et le programme peut se bloquer temporairement."
                                    Case "PTB", "PTG"
                                        Label5.Text = "Obter informações do ficheiro do controlador " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Isto pode demorar algum tempo e o programa pode congelar temporariamente"
                                    Case "ITA"
                                        Label5.Text = "Ottenere informazioni dal file del driver " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Questa operazione potrebbe richiedere del tempo e il programma potrebbe bloccarsi temporaneamente"
                                End Select
                            Case 1
                                Label5.Text = "Getting information from driver file " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "This may take some time and the program may temporarily freeze"
                            Case 2
                                Label5.Text = "Obteniendo información del archivo de controlador " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Esto puede llevar algo de tiempo y el programa podría congelarse temporalmente"
                            Case 3
                                Label5.Text = "Obtention des informations du fichier pilote " & Quote & Path.GetFileName(drvFile) & Quote & " en cours..." & CrLf & "Cette opération peut prendre un certain temps et le programme peut se bloquer temporairement."
                            Case 4
                                Label5.Text = "Obter informações do ficheiro do controlador " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Isto pode demorar algum tempo e o programa pode congelar temporariamente"
                            Case 5
                                Label5.Text = "Ottenere informazioni dal file del driver " & Quote & Path.GetFileName(drvFile) & Quote & "..." & CrLf & "Questa operazione potrebbe richiedere del tempo e il programma potrebbe bloccarsi temporaneamente"
                        End Select
                        Application.DoEvents()
                        Dim drvInfoCollection As DismDriverCollection = DismApi.GetDriverInfo(imgSession, drvFile)
                        If drvInfoCollection.Count > 0 Then DriverInfoList.Add(drvInfoCollection)
                    End If
                Next
            End Using
        Catch ex As Exception
            ' Cancel it
        Finally
            Try
                DismApi.Shutdown()
            Catch ex As Exception

            End Try
        End Try
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        Label5.Text = "Ready"
                    Case "ESN"
                        Label5.Text = "Listo"
                    Case "FRA"
                        Label5.Text = "Prêt"
                    Case "PTB", "PTG"
                        Label5.Text = "Pronto"
                    Case "ITA"
                        Label5.Text = "Pronto"
                End Select
            Case 1
                Label5.Text = "Ready"
            Case 2
                Label5.Text = "Listo"
            Case 3
                Label5.Text = "Prêt"
            Case 4
                Label5.Text = "Pronto"
            Case 5
                Label5.Text = "Pronto"
        End Select
    End Sub

    Sub DisplayDriverInformation(HWTarget As Integer)
        Dim CurrentDriverCollection As DismDriverCollection = DriverInfoList(ListBox1.SelectedIndex)
        For Each DriverPackageInfo As DismDriver In CurrentDriverCollection
            If CurrentDriverCollection.IndexOf(DriverPackageInfo) = HWTarget - 1 Then
                Label9.Text = DriverPackageInfo.HardwareDescription
                Label11.Text = DriverPackageInfo.HardwareId
                Label14.Text = DriverPackageInfo.CompatibleIds
                Label15.Text = DriverPackageInfo.ExcludeIds
                Label18.Text = DriverPackageInfo.ManufacturerName
                Label19.Text = Casters.CastDismArchitecture(DriverPackageInfo.Architecture, True)
                If Label14.Text = "" Then
                    Select Case MainForm.Language
                        Case 0
                            Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                                Case "ENU", "ENG"
                                    Label14.Text = "None declared by the hardware manufacturer"
                                Case "ESN"
                                    Label14.Text = "Ninguno declarado por el fabricante del hardware"
                                Case "FRA"
                                    Label14.Text = "Aucune déclarée par le fabricant du matériel"
                                Case "PTB", "PTG"
                                    Label14.Text = "Nenhum declarado pelo fabricante do hardware"
                                Case "ITA"
                                    Label14.Text = "Nessuno dichiarato dal produttore dell'hardware"
                            End Select
                        Case 1
                            Label14.Text = "None declared by the hardware manufacturer"
                        Case 2
                            Label14.Text = "Ninguno declarado por el fabricante del hardware"
                        Case 3
                            Label14.Text = "Aucune déclarée par le fabricant du matériel"
                        Case 4
                            Label14.Text = "Nenhum declarado pelo fabricante do hardware"
                        Case 5
                            Label14.Text = "Nessuno dichiarato dal produttore dell'hardware"
                    End Select
                End If
                If Label15.Text = "" Then
                    Select Case MainForm.Language
                        Case 0
                            Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                                Case "ENU", "ENG"
                                    Label15.Text = "None declared by the hardware manufacturer"
                                Case "ESN"
                                    Label15.Text = "Ninguno declarado por el fabricante del hardware"
                                Case "FRA"
                                    Label15.Text = "Aucune déclarée par le fabricant du matériel"
                                Case "PTB", "PTG"
                                    Label15.Text = "Nenhum declarado pelo fabricante do hardware"
                                Case "ITA"
                                    Label15.Text = "Nessuno dichiarato dal produttore dell'hardware"
                            End Select
                        Case 1
                            Label15.Text = "None declared by the hardware manufacturer"
                        Case 2
                            Label15.Text = "Ninguno declarado por el fabricante del hardware"
                        Case 3
                            Label15.Text = "Aucune déclarée par le fabricant du matériel"
                        Case 4
                            Label15.Text = "Nenhum declarado pelo fabricante do hardware"
                        Case 5
                            Label15.Text = "Nessuno dichiarato dal produttore dell'hardware"
                    End Select
                End If
                Exit For
            End If
        Next
    End Sub

    Sub DisplayHardwareTargetOverview()
        ' This function is called when the user clicks on the "Jump to target" button
        If ListBox1.SelectedItems.Count <> 1 Then
            ' Don't continue
            Exit Sub
        Else
            JumpTo = -1
            ComboBox1.Text = ""
            Dim CurrentDriverCollection As DismDriverCollection = DriverInfoList(ListBox1.SelectedIndex)
            For Each DriverPackageInfo As DismDriver In CurrentDriverCollection
                ComboBox1.Items.Add(CurrentDriverCollection.IndexOf(DriverPackageInfo) + 1 & " - " & DriverPackageInfo.HardwareDescription & " (" & DriverPackageInfo.HardwareId & ")")
            Next
        End If
    End Sub

    Private Sub ListBox1_DragEnter(sender As Object, e As DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub ListBox1_DragDrop(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
        Dim PackageFiles() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each PackageFile In PackageFiles
            If Path.GetExtension(PackageFile).EndsWith("inf", StringComparison.OrdinalIgnoreCase) Then
                ListBox1.Items.Add(PackageFile)
            End If
        Next
        Button3.Enabled = True
        Button8.Enabled = True
        GetDriverInformation()
    End Sub

    Private Sub GetDriverInfo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not MainForm.MountedImageDetectorBW.IsBusy Then Call MainForm.MountedImageDetectorBW.RunWorkerAsync()
        MainForm.WatcherTimer.Enabled = True
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            If ListBox1.SelectedItems.Count = 1 Then
                JumpToPanel.Visible = False
                NoDrvPanel.Visible = False
                DrvPackageInfoPanel.Visible = True
                Button2.Enabled = True
                If Not CurrentHWFile = ListBox1.SelectedIndex Then
                    Select Case MainForm.Language
                        Case 0
                            Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                                Case "ENU", "ENG"
                                    Label7.Text = "Hardware target 1 of " & DriverInfoList(ListBox1.SelectedIndex).Count
                                Case "ESN"
                                    Label7.Text = "Hardware de destino 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                                Case "FRA"
                                    Label7.Text = "Cible matérielle 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                                Case "PTB", "PTG"
                                    Label7.Text = "Equipamento-alvo 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                                Case "ITA"
                                    Label7.Text = "Destinazione hardware 1 di " & DriverInfoList(ListBox1.SelectedIndex).Count
                            End Select
                        Case 1
                            Label7.Text = "Hardware target 1 of " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case 2
                            Label7.Text = "Hardware de destino 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case 3
                            Label7.Text = "Cible matérielle 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case 4
                            Label7.Text = "Equipamento-alvo 1 de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case 5
                            Label7.Text = "Destinazione hardware 1 di " & DriverInfoList(ListBox1.SelectedIndex).Count
                    End Select
                End If
                If Not CurrentHWFile = ListBox1.SelectedIndex Then CurrentHWTarget = 1
                Button4.Enabled = False
                Button5.Enabled = True
                If Not CurrentHWFile = ListBox1.SelectedIndex Then DisplayDriverInformation(1)
                CurrentHWFile = ListBox1.SelectedIndex
                Button9.Enabled = True
            Else
                NoDrvPanel.Visible = True
                DrvPackageInfoPanel.Visible = False
                Button2.Enabled = False
                Button9.Enabled = False
            End If
        Catch ex As Exception
            ListBox1.Items.Remove(ListBox1.SelectedItem)
            NoDrvPanel.Visible = True
            DrvPackageInfoPanel.Visible = False
            If ListBox1.Items.Count < 1 Then
                Button2.Enabled = False
                Button3.Enabled = False
                Button8.Enabled = False
                Button9.Enabled = False
            End If
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DriverInfoList.RemoveAt(ListBox1.SelectedIndex)
        ListBox1.Items.Remove(ListBox1.SelectedItem)
        If ListBox1.Items.Count >= 1 Then
            Button2.Enabled = True
            Button3.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
        Else
            Button2.Enabled = False
            Button3.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
        End If
        NoDrvPanel.Visible = True
        DrvPackageInfoPanel.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DriverInfoList.Clear()
        ListBox1.Items.Clear()
        Button2.Enabled = False
        Button3.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        NoDrvPanel.Visible = True
        DrvPackageInfoPanel.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If CurrentHWTarget > 1 Then
            DisplayDriverInformation(CurrentHWTarget - 1)
            CurrentHWTarget -= 1
            Select Case MainForm.Language
                Case 0
                    Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                        Case "ENU", "ENG"
                            Label7.Text = "Hardware target " & CurrentHWTarget & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "ESN"
                            Label7.Text = "Hardware de destino " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "FRA"
                            Label7.Text = "Cible matérielle " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "PTB", "PTG"
                            Label7.Text = "Equipamento-alvo " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "ITA"
                            Label7.Text = "Destinazione hardware " & CurrentHWTarget & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
                    End Select
                Case 1
                    Label7.Text = "Hardware target " & CurrentHWTarget & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 2
                    Label7.Text = "Hardware de destino " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 3
                    Label7.Text = "Cible matérielle " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 4
                    Label7.Text = "Equipamento-alvo " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 5
                    Label7.Text = "Destinazione hardware " & CurrentHWTarget & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
            End Select
            Button5.Enabled = True
            If CurrentHWTarget = 1 Then Button4.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If CurrentHWTarget < DriverInfoList(ListBox1.SelectedIndex).Count Then
            DisplayDriverInformation(CurrentHWTarget + 1)
            CurrentHWTarget += 1
            Select Case MainForm.Language
                Case 0
                    Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                        Case "ENU", "ENG"
                            Label7.Text = "Hardware target " & CurrentHWTarget & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "ESN"
                            Label7.Text = "Hardware de destino " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "FRA"
                            Label7.Text = "Cible matérielle " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "PTB", "PTG"
                            Label7.Text = "Equipamento-alvo " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                        Case "ITA"
                            Label7.Text = "Destinazione hardware " & CurrentHWTarget & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
                    End Select
                Case 1
                    Label7.Text = "Hardware target " & CurrentHWTarget & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 2
                    Label7.Text = "Hardware de destino " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 3
                    Label7.Text = "Cible matérielle " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 4
                    Label7.Text = "Equipamento-alvo " & CurrentHWTarget & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                Case 5
                    Label7.Text = "Destinazione hardware " & CurrentHWTarget & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
            End Select
            Button4.Enabled = True
            If CurrentHWTarget = DriverInfoList(ListBox1.SelectedIndex).Count Then Button5.Enabled = False
        End If
    End Sub

    Private Sub Button4_MouseHover(sender As Object, e As EventArgs) Handles Button4.MouseHover
        Dim msg As String = ""
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        msg = "Previous hardware target"
                    Case "ESN"
                        msg = "Anterior hardware de destino"
                    Case "FRA"
                        msg = "Cible matérielle précédente"
                    Case "PTB", "PTG"
                        msg = "Equipamento-alvo anterior"
                    Case "ITA"
                        msg = "Destinazione hardware precedente"
                End Select
            Case 1
                msg = "Previous hardware target"
            Case 2
                msg = "Anterior hardware de destino"
            Case 3
                msg = "Cible matérielle précédente"
            Case 4
                msg = "Equipamento-alvo anterior"
            Case 5
                msg = "Destinazione hardware precedente"
        End Select
        ButtonTT.SetToolTip(sender, msg)
    End Sub

    Private Sub Button5_MouseHover(sender As Object, e As EventArgs) Handles Button5.MouseHover
        Dim msg As String = ""
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        msg = "Next hardware target"
                    Case "ESN"
                        msg = "Siguiente hardware de destino"
                    Case "FRA"
                        msg = "Prochaine cible matérielle"
                    Case "PTB", "PTG"
                        msg = "Equipamento-alvo seguinte"
                    Case "ITA"
                        msg = "Prossima destinazione hardware"
                End Select
            Case 1
                msg = "Next hardware target"
            Case 2
                msg = "Siguiente hardware de destino"
            Case 3
                msg = "Prochaine cible matérielle"
            Case 4
                msg = "Equipamento-alvo seguinte"
            Case 5
                msg = "Prossima destinazione hardware"
        End Select
        ButtonTT.SetToolTip(sender, msg)
    End Sub

    Private Sub Button6_MouseHover(sender As Object, e As EventArgs) Handles Button6.MouseHover
        Dim msg As String = ""
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        msg = "Jump to specific hardware target"
                    Case "ESN"
                        msg = "Saltar a hardware de destino específico"
                    Case "FRA"
                        msg = "Sauter à la cible matérielle spécifique"
                    Case "PTB", "PTG"
                        msg = "Saltar para um equipamento-alvo específico"
                    Case "ITA"
                        msg = "Salta a una destinazione hardware specifica"
                End Select
            Case 1
                msg = "Jump to specific hardware target"
            Case 2
                msg = "Saltar a hardware de destino específico"
            Case 3
                msg = "Sauter à la cible matérielle spécifique"
            Case 4
                msg = "Saltar para um equipamento-alvo específico"
            Case 5
                msg = "Salta a una destinazione hardware specifica"
        End Select
        ButtonTT.SetToolTip(sender, msg)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        JumpTo = ComboBox1.SelectedIndex + 1
        If JumpTo < 1 Then Exit Sub
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        Label7.Text = "Hardware target " & JumpTo & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
                    Case "ESN"
                        Label7.Text = "Hardware de destino " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                    Case "FRA"
                        Label7.Text = "Cible matérielle " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                    Case "PTB", "PTG"
                        Label7.Text = "Equipamento-alvo " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
                    Case "ITA"
                        Label7.Text = "Destinazione hardware " & JumpTo & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
                End Select
            Case 1
                Label7.Text = "Hardware target " & JumpTo & " of " & DriverInfoList(ListBox1.SelectedIndex).Count
            Case 2
                Label7.Text = "Hardware de destino " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
            Case 3
                Label7.Text = "Cible matérielle " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
            Case 4
                Label7.Text = "Equipamento-alvo " & JumpTo & " de " & DriverInfoList(ListBox1.SelectedIndex).Count
            Case 5
                Label7.Text = "Destinazione hardware " & JumpTo & " di " & DriverInfoList(ListBox1.SelectedIndex).Count
        End Select
        CurrentHWTarget = JumpTo
        DisplayDriverInformation(JumpTo)
        JumpToPanel.Visible = False
        If CurrentHWTarget = DriverInfoList(ListBox1.SelectedIndex).Count Then Button5.Enabled = False
        If CurrentHWTarget = 1 Then Button4.Enabled = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        JumpToPanel.Visible = True
        Button4.Enabled = True
        Button5.Enabled = True
        ComboBox1.Items.Clear()
        DisplayHardwareTargetOverview()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        Try
            If ListView1.SelectedItems.Count = 1 Then
                Panel4.Visible = True
                Panel7.Visible = False
                Dim drv As DismDriverPackage = Nothing
                If SearchBox1.Text = "" Then
                    drv = InstalledDriverList(ListView1.FocusedItem.Index)
                Else
                    drv = SearchedDriverList(ListView1.FocusedItem.Index)
                End If
                Label23.Text = drv.PublishedName
                Label25.Text = Path.GetFileName(drv.OriginalFileName)
                Select Case MainForm.Language
                    Case 0
                        Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                            Case "ENU", "ENG"
                                Label27.Text = If(drv.BootCritical, "Yes", "No")
                                Label34.Text = If(drv.InBox, "Yes", "No")
                            Case "ESN"
                                Label27.Text = If(drv.BootCritical, "Sí", "No")
                                Label34.Text = If(drv.InBox, "Sí", "No")
                            Case "FRA"
                                Label27.Text = If(drv.BootCritical, "Oui", "Non")
                                Label34.Text = If(drv.InBox, "Oui", "Non")
                            Case "PTB", "PTG"
                                Label27.Text = If(drv.BootCritical, "Sim", "Não")
                                Label34.Text = If(drv.InBox, "Sim", "Não")
                            Case "ITA"
                                Label27.Text = If(drv.BootCritical, "Sì", "No")
                                Label34.Text = If(drv.InBox, "Sì", "No")
                        End Select
                    Case 1
                        Label27.Text = If(drv.BootCritical, "Yes", "No")
                        Label34.Text = If(drv.InBox, "Yes", "No")
                    Case 2
                        Label27.Text = If(drv.BootCritical, "Sí", "No")
                        Label34.Text = If(drv.InBox, "Sí", "No")
                    Case 3
                        Label27.Text = If(drv.BootCritical, "Oui", "Non")
                        Label34.Text = If(drv.InBox, "Oui", "Non")
                    Case 4
                        Label27.Text = If(drv.BootCritical, "Sim", "Não")
                        Label34.Text = If(drv.InBox, "Sim", "Não")
                    Case 5
                        Label27.Text = If(drv.BootCritical, "Sì", "No")
                        Label34.Text = If(drv.InBox, "Sì", "No")
                End Select
                Label29.Text = drv.Version.ToString()
                Label32.Text = drv.ClassName
                Label35.Text = drv.ProviderName
                Label38.Text = drv.Date
                Label40.Text = drv.ClassDescription
                Label42.Text = drv.ClassGuid
                Label44.Text = Casters.CastDismSignatureStatus(drv.DriverSignature, True)
                Label46.Text = drv.CatalogFile
                Dim signer As String = DriverSignerViewer.GetSignerInfo(drv.OriginalFileName)
                If Not (signer Is Nothing OrElse signer = "") Then
                    Debug.WriteLine("Driver file: {0} ; Signer: {1}", Path.GetFileName(drv.OriginalFileName), signer)
                    Select Case MainForm.Language
                        Case 0
                            Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                                Case "ENU", "ENG"
                                    Label44.Text &= " by " & signer
                                Case "ESN"
                                    Label44.Text &= " por " & signer
                                Case "FRA"
                                    Label44.Text &= " par " & signer
                                Case "PTB", "PTG"
                                    Label44.Text &= " por " & signer
                                Case "ITA"
                                    Label44.Text &= " da " & signer
                            End Select
                        Case 1
                            Label44.Text &= " by " & signer
                        Case 2
                            Label44.Text &= " por " & signer
                        Case 3
                            Label44.Text &= " par " & signer
                        Case 4
                            Label44.Text &= " por " & signer
                        Case 5
                            Label44.Text &= " da " & signer
                    End Select
                End If
            Else
                Panel4.Visible = False
                Panel7.Visible = True
            End If
        Catch ex As Exception
            Panel4.Visible = False
            Panel7.Visible = True
        End Try
    End Sub

    Private Sub Label25_MouseHover(sender As Object, e As EventArgs) Handles Label25.MouseHover
        ButtonTT.SetToolTip(sender, InstalledDriverList(ListView1.FocusedItem.Index).OriginalFileName)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Visible = False
        BGProcsAdvSettings.ShowDialog(MainForm)
        If BGProcsAdvSettings.DialogResult = Windows.Forms.DialogResult.OK And BGProcsAdvSettings.NeedsDriverChecks Then Close() Else Visible = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If MainForm.ImgInfoSFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If Not ImgInfoSaveDlg.IsDisposed Then ImgInfoSaveDlg.Dispose()
            If ImgInfoSaveDlg.DriverPkgs.Count > 0 Then ImgInfoSaveDlg.DriverPkgs.Clear()
            ImgInfoSaveDlg.SourceImage = MainForm.SourceImg
            ImgInfoSaveDlg.SaveTarget = MainForm.ImgInfoSFD.FileName
            ImgInfoSaveDlg.ImgMountDir = If(Not MainForm.OnlineManagement, MainForm.MountDir, "")
            ImgInfoSaveDlg.OnlineMode = MainForm.OnlineManagement
            ImgInfoSaveDlg.OfflineMode = MainForm.OfflineManagement
            ImgInfoSaveDlg.AllDrivers = MainForm.AllDrivers
            ImgInfoSaveDlg.SkipQuestions = MainForm.SkipQuestions
            ImgInfoSaveDlg.AutoCompleteInfo = MainForm.AutoCompleteInfo
            ImgInfoSaveDlg.ForceAppxApi = False
            ImgInfoSaveDlg.SaveTask = If(InfoFromDrvPackagesPanel.Visible, 8, 7)
            If InfoFromDrvPackagesPanel.Visible Then
                For Each drvFile In ListBox1.Items
                    If File.Exists(drvFile) Then ImgInfoSaveDlg.DriverPkgs.Add(drvFile)
                Next
            End If
            ImgInfoSaveDlg.ShowDialog()
            InfoSaveResults.Show()
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DriverFileInfoDlg.ShowDialog()
    End Sub

    Sub SearchDrivers(sQuery As String, OriginalNames As Boolean)
        If InstalledDriverInfo.Count > 0 Then
            For Each InstalledDriver As DismDriverPackage In InstalledDriverInfo
                If OriginalNames Then
                    If (Path.GetFileName(InstalledDriver.OriginalFileName)).ToLower().Contains(sQuery.Replace("og:", "").Trim().ToLower()) Then
                        ListView1.Items.Add(New ListViewItem(New String() {InstalledDriver.PublishedName, Path.GetFileName(InstalledDriver.OriginalFileName)}))
                        SearchedDriverList.Add(InstalledDriver)
                    End If
                Else
                    If InstalledDriver.PublishedName.ToLower().Contains(sQuery.ToLower()) Then
                        ListView1.Items.Add(New ListViewItem(New String() {InstalledDriver.PublishedName, Path.GetFileName(InstalledDriver.OriginalFileName)}))
                        SearchedDriverList.Add(InstalledDriver)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub SearchBox1_TextChanged(sender As Object, e As EventArgs) Handles SearchBox1.TextChanged
        ListView1.Items.Clear()
        SearchedDriverList.Clear()
        If SearchBox1.Text <> "" Then
            SearchDrivers(SearchBox1.Text, SearchBox1.Text.StartsWith("og:", StringComparison.OrdinalIgnoreCase))
        Else
            For Each InstalledDriver As DismDriverPackage In InstalledDriverInfo
                ListView1.Items.Add(New ListViewItem(New String() {InstalledDriver.PublishedName, Path.GetFileName(InstalledDriver.OriginalFileName)}))
            Next
        End If
    End Sub
End Class
