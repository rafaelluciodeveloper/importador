unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, DB, DBTables, ComCtrls, Buttons, JvZlibMultiple, FileCtrl,
  JvComponentBase, IniFiles, TlHelp32, ExtCtrls;

type
    TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    TabelaDados: TStringGrid;
    Label1: TLabel;
    EditCaminhoContabil: TEdit;
    BtnEscolherArquivo: TBitBtn;
    BtnImportar: TBitBtn;
    ProgressBar1: TProgressBar;
    ComboBoxEmpresas: TComboBox;
    Label2: TLabel;
    BtnCarregarEmpresas: TBitBtn;
    Label3: TLabel;
    LabelCaminhoEmpresa: TLabel;
    JvZlibMultiple1: TJvZlibMultiple;
    procedure BtnCarregarEmpresasClick(Sender: TObject);
    procedure BtnEscolherArquivoClick(Sender: TObject);
    procedure ComboBoxEmpresasChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure BtnImportarClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BtnCarregarEmpresasClick(Sender: TObject);
var
  Tabela: TTable;
  Caminho, Codigo, Descricao: string;
begin
  ComboBoxEmpresas.Items.Clear;

  Caminho := Trim(EditCaminhoContabil.Text);
  if Caminho = '' then
  begin
    MessageDlg('Informe o caminho da empresa no "Caminho Contábil".', mtWarning, [mbOK], 0);
    Exit;
  end;

  if not FileExists(IncludeTrailingPathDelimiter(Caminho) + 'Empresas.db') then
  begin
    MessageDlg('Tabela Empresas.db não encontrada no caminho informado.', mtWarning, [mbOK], 0);
    Exit;
  end;

  Tabela := TTable.Create(nil);
  try
    Tabela.DatabaseName := Caminho;
    Tabela.TableName := 'Empresas.db';
    Tabela.TableType := ttParadox;
    Tabela.Exclusive := False;
    Tabela.ReadOnly := True;

    try
      Tabela.Open;
    except
      on E: Exception do
      begin
        MessageDlg('Erro ao abrir a tabela Empresas.db:' + sLineBreak + E.Message , mtError, [mbOK], 0);
        Exit;
      end;
    end;

    if not Tabela.Active then
    begin
      MessageDlg('Não foi possível abrir a tabela Empresas.db.', mtError, [mbOK], 0);
      Exit;
    end;

    Tabela.First;
    while not Tabela.Eof do
    begin
      Codigo := Tabela.FieldByName('Codigo').AsString;
      Descricao := Tabela.FieldByName('RazaoSocial').AsString;
      ComboBoxEmpresas.Items.Add(Codigo + ' - ' + Descricao);
      Tabela.Next;
    end;

    if ComboBoxEmpresas.Items.Count > 0 then
      ComboBoxEmpresas.ItemIndex := 0;
      if ComboBoxEmpresas.ItemIndex >= 0 then
      begin
        Codigo := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex], 1, Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) - 1);
        Codigo := FormatFloat('000', StrToIntDef(Codigo, 0));
        LabelCaminhoEmpresa.Caption := IncludeTrailingPathDelimiter(Trim(EditCaminhoContabil.Text)) + 'Emp' + Codigo;
      end;
    BtnEscolherArquivo.Enabled := True;
    ComboBoxEmpresas.Enabled := True;

  finally
    Tabela.Free;
  end;
end;

procedure TForm1.BtnEscolherArquivoClick(Sender: TObject);
var
  CSVFile: TextFile;
  Line, Field: string;
  Row, Col, P, StartPos, CampoCount: Integer;
  Fields: array[0..5] of string;
  GridWidth, LarguraRestante: Integer;
begin
  if OpenDialog1.Execute then
  begin
    // Sempre limpa o grid antes de qualquer coisa
    TabelaDados.RowCount := 2;
    TabelaDados.ColCount := 6;
    TabelaDados.FixedRows := 1;

    for Row := 1 to TabelaDados.RowCount - 1 do
      for Col := 0 to TabelaDados.ColCount - 1 do
        TabelaDados.Cells[Col, Row] := '';

    // Cabeçalhos
    TabelaDados.Cells[0, 0] := 'Data';
    TabelaDados.Cells[1, 0] := 'Débito';
    TabelaDados.Cells[2, 0] := 'Crédito';
    TabelaDados.Cells[3, 0] := 'Histórico Padrão';
    TabelaDados.Cells[4, 0] := 'Complemento Histórico';
    TabelaDados.Cells[5, 0] := 'Valor';

    AssignFile(CSVFile, OpenDialog1.FileName);
    Reset(CSVFile);

    Row := 1;

    while not Eof(CSVFile) do
    begin
      ReadLn(CSVFile, Line);

      // Inicializa os campos
      for Col := 0 to 5 do
        Fields[Col] := '';

      // Parsing manual preservando campos em branco
      StartPos := 1;
      CampoCount := 0;

      for Col := 0 to 5 do
      begin
        P := Pos(';', Copy(Line, StartPos, Length(Line) - StartPos + 1));
        if P > 0 then
        begin
          Field := Copy(Line, StartPos, P - 1);
          Fields[Col] := Trim(Field);
          StartPos := StartPos + P;
          Inc(CampoCount);
        end
        else
        begin
          Fields[Col] := Trim(Copy(Line, StartPos, MaxInt));
          Inc(CampoCount);
          Break;
        end;
      end;

      // Validação do layout
      if CampoCount < 6 then
      begin
        CloseFile(CSVFile);

        // Limpa tudo novamente
        TabelaDados.RowCount := 2;
        TabelaDados.ColCount := 6;
        TabelaDados.FixedRows := 1;

        for Row := 1 to TabelaDados.RowCount - 1 do
          for Col := 0 to TabelaDados.ColCount - 1 do
            TabelaDados.Cells[Col, Row] := '';

        for Col := 0 to 5 do
          case Col of
            0: TabelaDados.Cells[Col, 0] := 'Data';
            1: TabelaDados.Cells[Col, 0] := 'Débito';
            2: TabelaDados.Cells[Col, 0] := 'Crédito';
            3: TabelaDados.Cells[Col, 0] := 'Histórico Padrão';
            4: TabelaDados.Cells[Col, 0] := 'Complemento Histórico';
            5: TabelaDados.Cells[Col, 0] := 'Valor';
          end;

        MessageDlg('O layout do arquivo está incorreto na linha ' + IntToStr(Row), mtWarning , [mbOK], 0);
        BtnImportar.Enabled := False;
        TabelaDados.Enabled := False;
        Exit;
      end;

      // Preenche os dados no grid
      TabelaDados.RowCount := Row + 1;
      for Col := 0 to 5 do
        TabelaDados.Cells[Col, Row] := Fields[Col];

      Inc(Row);
    end;

    CloseFile(CSVFile);

    // Ajuste das larguras
    GridWidth := TabelaDados.ClientWidth;

    TabelaDados.ColWidths[0] := 80;   // Data
    TabelaDados.ColWidths[1] := 80;   // Débito
    TabelaDados.ColWidths[2] := 80;   // Crédito
    TabelaDados.ColWidths[3] := 120;  // Histórico Padrão
    TabelaDados.ColWidths[5] := 80;   // Valor

    LarguraRestante := GridWidth
      - (TabelaDados.ColWidths[0] + TabelaDados.ColWidths[1] +
         TabelaDados.ColWidths[2] + TabelaDados.ColWidths[3] +
         TabelaDados.ColWidths[5]);

    if LarguraRestante > 100 then
      TabelaDados.ColWidths[4] := LarguraRestante
    else
      TabelaDados.ColWidths[4] := 100;
    BtnImportar.Enabled := True;
    TabelaDados.Enabled := True;
  end;
end;
procedure TForm1.ComboBoxEmpresasChange(Sender: TObject);
var
  Codigo: string;
  Row, Col: Integer;
begin
  if ComboBoxEmpresas.ItemIndex >= 0 then
  begin
    Codigo := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex], 1, Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) - 1);
    Codigo := FormatFloat('000', StrToIntDef(Codigo, 0));
    LabelCaminhoEmpresa.Caption := IncludeTrailingPathDelimiter(Trim(EditCaminhoContabil.Text)) + 'Emp' + Codigo;
    TabelaDados.RowCount := 2;
    TabelaDados.ColCount := 6;
    TabelaDados.FixedRows := 1;
    for Row := 1 to TabelaDados.RowCount - 1 do
      for Col := 0 to TabelaDados.ColCount - 1 do
        TabelaDados.Cells[Col, Row] := '';

    for Col := 0 to 5 do
      case Col of
        0: TabelaDados.Cells[Col, 0] := 'Data';
        1: TabelaDados.Cells[Col, 0] := 'Débito';
        2: TabelaDados.Cells[Col, 0] := 'Crédito';
        3: TabelaDados.Cells[Col, 0] := 'Histórico Padrão';
        4: TabelaDados.Cells[Col, 0] := 'Complemento Histórico';
        5: TabelaDados.Cells[Col, 0] := 'Valor';
      end;
  end;
end;

procedure TForm1.FormShow(Sender: TObject);
var
  Row, Col: Integer;
  INI: TIniFile;
  INIPath, CaminhoConfig: string;
  Snap: THandle;
  ProcEntry: TProcessEntry32;
  Encontrado: Boolean;
begin
  Encontrado := False;
  Snap := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if Snap <> INVALID_HANDLE_VALUE then
  begin
    ProcEntry.dwSize := SizeOf(TProcessEntry32);
    if Process32First(Snap, ProcEntry) then
    begin
      repeat
        if SameText(ExtractFileName(ProcEntry.szExeFile), 'Contabil.exe') then
        begin
          Encontrado := True;
          Break;
        end;
      until not Process32Next(Snap, ProcEntry);
    end;
    CloseHandle(Snap);
  end;

  if Encontrado then
  begin
    MessageDlg('O sistema Contabil.exe está em execução.' + sLineBreak +
                'Feche-o antes de continuar.', mtWarning, [mbOK],0);
    Application.Terminate;
    Exit;
  end;

  INIPath := ExtractFilePath(Application.ExeName) + 'Config.ini';
  if FileExists(INIPath) then
  begin
    INI := TIniFile.Create(INIPath);
    try
      CaminhoConfig := INI.ReadString('Configuracoes', 'Caminho', '');
      if CaminhoConfig <> '' then
        EditCaminhoContabil.Text := CaminhoConfig;
    finally
      INI.Free;
    end;
  end
  else
  begin
    MessageDlg('Arquivo Config.ini não encontrado na pasta do sistema.', mtWarning , [mbOK] , 0);
    Application.Terminate;
  end;
  TabelaDados.RowCount := 2;
  TabelaDados.ColCount := 6;
  TabelaDados.FixedRows := 1;

  for Row := 1 to TabelaDados.RowCount - 1 do
    for Col := 0 to TabelaDados.ColCount - 1 do
      TabelaDados.Cells[Col, Row] := '';

  for Col := 0 to 5 do
    case Col of
      0: TabelaDados.Cells[Col, 0] := 'Data';
      1: TabelaDados.Cells[Col, 0] := 'Débito';
      2: TabelaDados.Cells[Col, 0] := 'Crédito';
      3: TabelaDados.Cells[Col, 0] := 'Histórico Padrão';
      4: TabelaDados.Cells[Col, 0] := 'Complemento Histórico';
      5: TabelaDados.Cells[Col, 0] := 'Valor';
    end;
    TabelaDados.ColWidths[0] := 80;   // Data
    TabelaDados.ColWidths[1] := 80;   // Débito
    TabelaDados.ColWidths[2] := 80;   // Crédito
    TabelaDados.ColWidths[3] := 100;
    TabelaDados.ColWidths[4] := 430;
    TabelaDados.ColWidths[5] := 80;   // Valor
end;

procedure TForm1.Button4Click(Sender: TObject);
var
  CaminhoOrigEmp, CaminhoDestEmp, CodigoEmpresa: string;
begin
  CaminhoOrigEmp := IncludeTrailingPathDelimiter(LabelCaminhoEmpresa.Caption);

  if not DirectoryExists(CaminhoOrigEmp) then
  begin
    MessageDlg('Pasta não encontrada: ' + CaminhoOrigEmp, mtWarning , [mbOK], 0);
    Exit;
  end;

  // Pega o código da empresa do ComboBox
  if ComboBoxEmpresas.ItemIndex < 0 then
  begin
    MessageDlg('Selecione uma empresa', mtInformation, [mbOK], 0);
    Exit;
  end;

  CodigoEmpresa := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex], 1,
    Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) - 1);
  CodigoEmpresa := FormatFloat('000', StrToIntDef(CodigoEmpresa, 0));

  CaminhoDestEmp := CaminhoOrigEmp + 'Emp' + CodigoEmpresa + '.bkp';

  // Zera progresso antes de começar
  ProgressBar1.Position := 0;
  ProgressBar1.Max := 0;
  Application.ProcessMessages;

  try
    // Compacta a pasta
    JvZlibMultiple1.CompressDirectory(CaminhoOrigEmp, False, CaminhoDestEmp); // False = não inclui subpastas

    if FileExists(CaminhoDestEmp) then
      MessageDlg('Compactado com sucesso: ' + CaminhoDestEmp ,mtInformation, [mbOK], 0)
    else
      MessageDlg('Erro: arquivo ZIP não foi criado.', mtError, [mbOK], 0);
  except
    on E: Exception do
      MessageDlg('Erro ao compactar: ' + E.Message , mtError, [mbOK], 0);
  end;

  ProgressBar1.Position := 0;
end;

procedure TForm1.BtnImportarClick(Sender: TObject);
var
  i, j: Integer;
  DataStr, ValorStr, DebitoStr, CreditoStr, HistoricoStr, ComplementoStr: string;
  QtDigitado: Integer;
  Vldigitado, ValorAtual , Valor: Double;
  TabelaLote, TabelaLanc: TTable;
  FormatSettings: TFormatSettings;
  DataList, SplitInfo, UltimosLancamentos: TStringList;
  Found: Boolean;
  Data: TDateTime;
  Debito, Credito, Historico, LancamentoAtual: Integer;
  DataChave: string;
  CaminhoEmp, CodigoEmpresa, RazaoSocial, CaminhoLotes, CaminhoLanc, CaminhoDestEmp: string;
begin
  // Confirma importacao
  if ComboBoxEmpresas.ItemIndex < 0 then
  begin
    MessageDlg('Selecione uma empresa', mtInformation, [mbOK], 0);
    Exit;
  end;

  CodigoEmpresa := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex], 1,
    Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) - 1);
  RazaoSocial := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex],
    Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) + 3, MaxInt);

  if MessageDlg('Deseja importar os lançamentos da empresa ' + CodigoEmpresa + ' - ' + RazaoSocial + '?',
    mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;

  // Configura decimal
  GetLocaleFormatSettings(LOCALE_SYSTEM_DEFAULT, FormatSettings);
  FormatSettings.DecimalSeparator := ',';

  CaminhoEmp := IncludeTrailingPathDelimiter(LabelCaminhoEmpresa.Caption);
  CaminhoLotes := CaminhoEmp + 'Lotes.db';
  CaminhoLanc := CaminhoEmp + 'Lancamentos.db';

  CodigoEmpresa := Copy(ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex], 1,
    Pos(' - ', ComboBoxEmpresas.Items[ComboBoxEmpresas.ItemIndex]) - 1);
  CodigoEmpresa := FormatFloat('000', StrToIntDef(CodigoEmpresa, 0));

  CaminhoDestEmp := CaminhoEmp + 'Emp' + CodigoEmpresa + '.bkp';

  // Faz backup antes de tudo
  JvZlibMultiple1.CompressDirectory(CaminhoEmp, False, CaminhoDestEmp);

  // Agrupar valores por data (para Lotes.db)
  DataList := TStringList.Create;
  DataList.Sorted := False;
  DataList.Duplicates := dupIgnore;

  for i := 1 to TabelaDados.RowCount - 1 do
  begin
    DataStr := Trim(TabelaDados.Cells[0, i]);
    ValorStr := Trim(TabelaDados.Cells[5, i]);
    if DataStr = '' then Continue;

    ValorAtual := StrToFloatDef(ValorStr, 0, FormatSettings);
    Found := False;

    for j := 0 to DataList.Count - 1 do
    begin
      if DataList.Names[j] = DataStr then
      begin
        SplitInfo := TStringList.Create;
        SplitInfo.Delimiter := '|';
        SplitInfo.DelimitedText := DataList.ValueFromIndex[j];

        Vldigitado := StrToFloatDef(SplitInfo[0], 0, FormatSettings);
        QtDigitado := StrToIntDef(SplitInfo[1], 0);

        Vldigitado := Vldigitado + ValorAtual;
        Inc(QtDigitado);

        DataList.ValueFromIndex[j] := FloatToStr(Vldigitado) + '|' + IntToStr(QtDigitado);
        SplitInfo.Free;
        Found := True;
        Break;
      end;
    end;

    if not Found then
      DataList.Add(DataStr + '=' + FloatToStr(ValorAtual) + '|1');
  end;

  ProgressBar1.Max := TabelaDados.RowCount - 1;
  ProgressBar1.Position := 0;

  // Inserir em Lotes.db
  TabelaLote := TTable.Create(nil);
  try
    TabelaLote.DatabaseName := CaminhoEmp;
    TabelaLote.TableName := 'Lotes.db';
    TabelaLote.TableType := ttParadox;
    TabelaLote.Open;

    for i := 0 to DataList.Count - 1 do
    begin
      DataStr := DataList.Names[i];
      ValorStr := DataList.ValueFromIndex[i];

      SplitInfo := TStringList.Create;
      SplitInfo.Delimiter := '|';
      SplitInfo.DelimitedText := ValorStr;

      Vldigitado := StrToFloatDef(SplitInfo[0], 0, FormatSettings);
      QtDigitado := StrToIntDef(SplitInfo[1], 0);
      SplitInfo.Free;

      TabelaLote.Append;
      TabelaLote.FieldByName('Data').AsString := DataStr;
      TabelaLote.FieldByName('Lote').AsInteger := 0;
      TabelaLote.FieldByName('Tipo').AsString := '0';
      TabelaLote.FieldByName('Vldigitado').AsFloat := Vldigitado;
      TabelaLote.FieldByName('QtDigitado').AsInteger := QtDigitado;
      TabelaLote.Post;
    end;
  finally
    TabelaLote.Free;
    DataList.Free;
  end;

  // Inserir em Lancamentos.db
  UltimosLancamentos := TStringList.Create;
  UltimosLancamentos.Sorted := False;
  UltimosLancamentos.Duplicates := dupIgnore;

  TabelaLanc := TTable.Create(nil);
  try
    TabelaLanc.DatabaseName := CaminhoEmp;
    TabelaLanc.TableName := 'Lancamentos.db';
    TabelaLanc.TableType := ttParadox;
    TabelaLanc.Open;

    for i := 1 to TabelaDados.RowCount - 1 do
    begin
      DataStr := Trim(TabelaDados.Cells[0, i]);
      DebitoStr := Trim(TabelaDados.Cells[1, i]);
      CreditoStr := Trim(TabelaDados.Cells[2, i]);
      HistoricoStr := Trim(TabelaDados.Cells[3, i]);
      ComplementoStr := Trim(TabelaDados.Cells[4, i]);
      ValorStr := Trim(TabelaDados.Cells[5, i]);

      if DataStr = '' then Continue;

      try
        Data := StrToDate(DataStr);
      except
        Continue;
      end;

      Debito := StrToIntDef(DebitoStr, 0);
      Credito := StrToIntDef(CreditoStr, 0);
      Historico := StrToIntDef(HistoricoStr, 0);
      Valor := StrToFloatDef(ValorStr, 0, FormatSettings);

      DataChave := DateToStr(Data);

      if UltimosLancamentos.IndexOfName(DataChave) = -1 then
      begin
        LancamentoAtual := 0;
        TabelaLanc.First;
        while not TabelaLanc.Eof do
        begin
          if Trunc(TabelaLanc.FieldByName('Data').AsDateTime) = Trunc(Data) then
          begin
            if TabelaLanc.FieldByName('Lancamento').AsInteger > LancamentoAtual then
              LancamentoAtual := TabelaLanc.FieldByName('Lancamento').AsInteger;
          end;
          TabelaLanc.Next;
        end;
        UltimosLancamentos.Add(DataChave + '=' + IntToStr(LancamentoAtual));
      end;

      LancamentoAtual := StrToInt(UltimosLancamentos.Values[DataChave]) + 1;
      UltimosLancamentos.Values[DataChave] := IntToStr(LancamentoAtual);

      TabelaLanc.Append;
      TabelaLanc.FieldByName('Data').AsDateTime := Data;
      TabelaLanc.FieldByName('Lote').AsInteger := 0;
      TabelaLanc.FieldByName('Lancamento').AsInteger := LancamentoAtual;
      TabelaLanc.FieldByName('Debito').AsInteger := Debito;
      TabelaLanc.FieldByName('Credito').AsInteger := Credito;
      TabelaLanc.FieldByName('Historico').AsInteger := Historico;
      TabelaLanc.FieldByName('Complemento').AsString := ComplementoStr;
      TabelaLanc.FieldByName('Valor').AsFloat := Valor;
      TabelaLanc.Post;

      ProgressBar1.Position := i;
      Application.ProcessMessages;
    end;
  finally
    UltimosLancamentos.Free;
    TabelaLanc.Free;
  end;

  ShowMessage('Importação de lançamentos concluída com sucesso!' + #13#10 +
  'Realizar procedimento de Atualizar saldo no contábil.' + #13#10 +
  'Processamento > Atualizar Saldo'
  );
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if MessageDlg('Deseja realmente sair do sistema?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    CanClose := True
  else
    CanClose := False;
end;
end.
