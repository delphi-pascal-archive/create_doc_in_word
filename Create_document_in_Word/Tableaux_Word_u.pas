unit Tableaux_Word_u;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Clipbrd, Buttons, DriveOleWord, ComCtrls,DateUtils,
 Grids, DBGrids,OleServer, variants,   WordXP,word2000;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Image1: TImage;
    Button2: TButton;
    OpenDialog1: TOpenDialog;
    Edit2: TEdit;
    SpeedButton4: TSpeedButton;
    Image2: TImage;
    Memo1: TMemo;
    Memo2: TMemo;
    Memo3: TMemo;
    Memo4: TMemo;
    SpeedButton6: TSpeedButton;
    Edit3: TEdit;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton10: TSpeedButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    Edit1: TEdit;
    Panel1: TPanel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    
  private
    { Private declarations }

  public

    { Public declarations }
  end;

var
  Form1: TForm1;
  wrdApp, wrdDoc: Variant;
    

implementation
 uses ComObj ;

{$R *.DFM}


procedure TForm1.Button1Click(Sender: TObject);
var
  StrToAdd   : String;
  ImageTempo : TPicture;
  mois : string ;
begin
   // image  logo ONERA
   ImageTempo := TPicture.Create;
   ImageTempo := Image1.Picture;
   mois := Format('%2.2d',[MonthOf(date)]);

   // Pour copier une image, je passe par le presse papier
   Clipboard.Assign(ImageTempo);

  // Creation d'une instance de Word visible
   CreerInstanceDeWord(wrdApp,true);
 // 
  // Creation d'un nouveau document
  CreerNouveauDocument(wrdDoc,wrdApp);
  ModeAffichage(wrdApp, 4);
  //----------- GENERATION DE L'ENTETE ---------------
  Activer_Entete(wrdapp);
  Alignement(wrdApp,Centre); // centrer
  CollerLePressePapier(wrdApp);
 // InsererTexte(wrdApp,'Example of management Word from Delphi');

  Activer_Corps_Document(wrdApp);

  // Police :
  TexteGras(wrdApp,true); // Texte en Gras
  TexteSouligne(wrdApp,FALSE); // Souligne le texte
  TexteTaille(wrdApp,12); // Taille de la police : 14
  Alignement(wrdApp,Centre); // centrer
  TexteFont(wrdApp,'Time'); // Choix de la police, on peut obtenir ces noms simplement
                    //  Comic Sans MS'  // avec un  if FontDialog1.Execute then Edit1.Text := FontDialog1.Font.Name;

   //----------- TITRE  ---------------
  // Insertion de Texte :            // #39
  InsererTexte(wrdApp,#10'Example of management Word from Delphi'#10); // #10 : retour a la ligne sinon, la police du dessous affecte la ligne

  //  fonte texte
  TexteGras(wrdApp,False); // Texte en Gras
  TexteSouligne(wrdApp,FALSE); // Souligne le texte
  TexteTaille(wrdApp,14); // Taille de la police : 20
  RetourLigne(wrdApp,1);
  Alignement(wrdApp,Droit); // centre
  RetourLigne(wrdApp,1);
  TexteFont(wrdApp,'Time'); // Choix de la police, on peut obtenir ces noms simplement

  // date
  InsererTexte(wrdApp,'Date, on ' +DateToStr(Date)+'.'+#10); // #10 : retour a la ligne sinon, la police du dessous affecte la ligne
  RetourLigne(wrdApp,1);
  // ref BE 'N° DMPH/L -xxx/'+mois+'/HM
   InsererTexte(wrdApp,'N° -xxx/'+mois+'/yy  '); // #10 : retour a la ligne sinon, la police du dessous affecte la ligne

   RetourLigne(wrdApp,1);
   //
   // image  logo DMPH     ImageTempo := TPicture.Create;
 
   ImageTempo := Image2.Picture;
   Alignement(wrdApp,Gauche);
   Clipboard.Assign(ImageTempo);
   CollerLePressePapier(wrdApp);  // logo DMPH

  // Police :
  TexteGras(wrdApp,false);   // Texte non Gras
  Alignement(wrdApp,Gauche); // alignement a gauche
  TexteSouligne(wrdApp,false); // texte non Souligne
  TexteTaille(wrdApp,12);    // Taille de la police : 12

   // ---------------- 1er tableau --------------------

  // CREATION DU TABLEAU DES ADRESSES
  CreerTableau(wrdApp,wrdDoc,1,2);

  // Choix des types de traits du tableau
  MiseEnFormeTableau(WrdDoc,1,None); // single

  // La case 4-4 : trait double
 // MiseEnFormeCelluleTableau(WrdDoc,1,4,4,Double,Double,Double,Double);
  //Choix des largeurs des cases
  // Choix des couleurs
  TableauCouleurColonne(wrdDoc,1,1,Jaune);
  TableauCouleurLigne(wrdDoc,1,1,GrisClair);
  TableauCouleurCase(wrdDoc,1,1,1,Blanc);

  // choix des alignements                Centre
 // TableauAlignementLigne(wrdDoc,1,1,Gauche);
 
  TableauAlignementCase(wrdDoc,1,1,1,Gauche); // N° TABLEAU cOL LIGNE

  TableauGrasLigne(wrdDoc,1,1,true);
  TableauGrasCase( wrdDoc,1,2,1,true);

  // remplissage des adresses
   TexteTaille(wrdApp,10); // Taille de la police : 20
   TexteFont(wrdApp,'Helvetica'); // Choix de la police, on peut obtenir ces noms simplement

   TableauTexteDansCase(wrdDoc,1,1,1,memo1.Text);
   TexteTaille(wrdApp,12); // Taille de la police : 20
   TexteFont(wrdApp,'Time'); // Choix de la police, on peut obtenir ces noms simplement
   TableauTexteDansCase(wrdDoc,1,2,1,memo2.Text);

  // Pour sortir du tableau
  AllerEnFinDuFichier(wrdApp);

// ------------------ BORDEREAU DE LIVRAISON -------------------------------------------

   // BORDEREAU DE LIVRAISON
  TexteGras(wrdApp,true); // Texte en Gras
  TexteTaille(wrdApp,14); // Taille de la police : 14
  Alignement(wrdApp,Centre); // centrer

   RetourLigne(wrdApp,1);
   InsererTexte(wrdApp,'TABLE N 2');  //
   RetourLigne(wrdApp,1);
   //
   // Réf
  TexteGras(wrdApp,false); // Texte en Gras
  TexteTaille(wrdApp,10); // Taille de la police : 10
  Alignement(wrdApp,Gauche); // centrer

  RetourLigne(wrdApp,1);
  InsererTexte(wrdApp,'Table');  //

 //   --------------------- 2 Ième  TABLEAU    --------------------------------

   //  2 Ième  TABLEAU DES DESIGNATIONS
   // CREATION DU TABLEAU  2 lignes 3 colonnes
    CreerTableau(wrdApp,wrdDoc,2,3);
   // Choix des types de traits du tableau  Triple
   MiseEnFormeTableau(WrdDoc,2, Single  );//     Emboss3D

   WrdDoc.Tables.Item(2).Cell(1,2).borders.item(2).LineWidth := 6;


  //Choix des largeurs des cases
  TableauLargeurCase(wrdDoc,2,1,315);
  TableauLargeurCase(wrdDoc,2,2,56);
  TableauLargeurCase(wrdDoc,2,3,160);

  // Choix des couleurs
   TableauCouleurLigne(wrdDoc,1,1,Blanc);
   TableauCouleurLigne(wrdDoc,2, 2 , rouge);

  // choix des alignements                Centre
 // TableauAlignementLigne(wrdDoc,1,1,Gauche);

    // choix des alignements        
    TableauAlignementCase(wrdDoc,2,1,1,Centre); // N° TABLEAU cOL LIGNE
    TableauAlignementCase(wrdDoc,2,2,1,Centre);
    TableauAlignementCase(wrdDoc,2,3,1,Centre);

  TableauGrasLigne(wrdDoc,1,1,true);
  TableauGrasCase( wrdDoc,2,1,1,true);
  TableauGrasCase( wrdDoc,2,2,1,true);
  TableauGrasCase( wrdDoc,2,3,1,true);
  // remplissage 1er ligne
  TableauTexteDansCase(wrdDoc,2,1,1,'COLUMN 1');
  TableauTexteDansCase(wrdDoc,2,2,1,'COLUMN 2');
  TableauTexteDansCase(wrdDoc,2,3,1,'COLUMN 3');

  // 2eme ligne
  //  TableauTexteDansCase(wrdDoc,1,1,2,'Toto');
 // TableauTexteDansCase(wrdDoc,1,2,2,'10');
  TableauGrasLigne(wrdDoc,1,1,False);
  TableauGrasCase( wrdDoc,2,1,1,False);
  TableauGrasCase( wrdDoc,2,2,1,False);
  TableauGrasCase( wrdDoc,2,3,1,False);
  // remplissage 1er ligne
    TableauAlignementCase(wrdDoc,2,1,2,Gauche); // N° TABLEAU cOL LIGNE
    TableauAlignementCase(wrdDoc,2,2,2,Centre);
    TableauAlignementCase(wrdDoc,2,3,2,Centre);

  TableauTexteDansCase(wrdDoc,2,1,2,Memo3.Text);

  TableauTexteDansCase(wrdDoc,2,2,2,#10#10' 1 ex.');

  TableauTexteDansCase(wrdDoc,2,3,2,Memo4.Text);
// --------------- fin 2ième tableau

 // curseur dans le tableau   // Var InstanceDeWord, Document
   TableauPlacerCurseurDansCase(wrdApp, wrdDoc  , 2, 2,2);
  // le déplacer
  TableauDeplacerCurseur(wrdApp,  BasDir, 2);


Panel1.Visible := true;
    left := 550  ;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin

     openDialog1.Filter :='Doc Word|*.doc';
     OpenDialog1.FileName := '*.doc';
     openDialog1.FilterIndex := 1;
     if OpenDialog1.Execute then
     begin
          CreerInstanceDeWord(wrdApp,true);
          OuvrirUnDocument(wrdApp,wrdDoc,OpenDialog1.FileName);
          edit2.Text := OpenDialog1.FileName;
     end;

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    left := screen.Width - width  ;
    edit1.Text := datetostr(date);
end;

procedure TForm1.SpeedButton4Click(Sender: TObject);
begin
     openDialog1.Filter :='Doc Word|*.doc';
     OpenDialog1.FileName := edit2.text ;

    if OpenDialog1.Execute then
     edit2.Text :=OpenDialog1.FileName
end;

procedure TForm1.SpeedButton6Click(Sender: TObject);
begin

      // début de ligne
      wrdapp.Selection.HomeKey ;// Unit:=wdStory
      // monte de n migne
      DeplacerCurseur( wrdApp,HautDir,100);
      // remplace le texte 'Châtillon' par celui de edit3.text
      wrdapp.Selection.Find.text := 'Date, on';
      wrdapp.Selection.Find.Replacement.text :=Edit3.Text;
      wrdapp.Selection.Find.Execute(replace := wdreplaceall);

end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if not VarIsEmpty(wrdDoc ) then
  begin
      FermerDocument(wrdDoc );
      wrdDoc := unAssigned ;
  end;
  if not VarIsEmpty(wrdApp) then
  begin

    FermerWord(wrdApp ) ;
    wrdApp := unAssigned ;
  end;
end;

procedure TForm1.SpeedButton7Click(Sender: TObject);

begin
  WrdDoc.Tables.Item(2).select;
  WrdDoc.Tables.Item(2).Cell(2,2).borders.item(2).LineStyle := Single;
   // largeurs des colonnes
  WrdDoc.Tables.Item(2).Columns.Width := 120;
end;

procedure TForm1.SpeedButton8Click(Sender: TObject);
begin
    close
end;

procedure TForm1.SpeedButton10Click(Sender: TObject);
begin
   Showmessage('Comment faire ?'#13'........');
   //WrdDoc.Tables.Item(2).Borders.InsideLineStyle :=  true ;
   //WrdDoc.selection.cells.borders.item(wdBorderLeft).LineWidth := WdLineWidth150pt;
end;

end.

