use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 86: Interface Property Accessors & Default Indexers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_interface_property_read_write() {
    let out = run_pascal(
        r#"
program Test;
type IData = interface
  ['{11111111-2222-3333-4444-555555555555}']
  function GetVal: Integer;
  procedure SetVal(v: Integer);
  property Val: Integer read GetVal write SetVal;
end;

type TDataImpl = class(TInterfacedObject, IData)
  private FVal: Integer;
  public
    function GetVal: Integer;
    procedure SetVal(v: Integer);
end;
function TDataImpl.GetVal: Integer; begin Result := FVal; end;
procedure TDataImpl.SetVal(v: Integer); begin FVal := v; end;

var d: IData;
begin
  d := TDataImpl.Create;
  d.Val := 100;
  WriteLn(d.Val);
end.
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_interface_property_read_only() {
    let out = run_pascal(
        r#"
program Test;
type IReadOnlyData = interface
  ['{22222222-3333-4444-5555-666666666666}']
  function GetCount: Integer;
  property Count: Integer read GetCount;
end;

type TCountImpl = class(TInterfacedObject, IReadOnlyData)
  public function GetCount: Integer;
end;
function TCountImpl.GetCount: Integer; begin Result := 42; end;

var c: IReadOnlyData;
begin
  c := TCountImpl.Create;
  WriteLn(c.Count);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_interface_default_indexer() {
    let out = run_pascal(
        r#"
program Test;
type IStringContainer = interface
  ['{33333333-4444-5555-6666-777777777777}']
  function GetItem(idx: Integer): String;
  property Items[idx: Integer]: String read GetItem; default;
end;

type TContainerImpl = class(TInterfacedObject, IStringContainer)
  public function GetItem(idx: Integer): String;
end;
function TContainerImpl.GetItem(idx: Integer): String;
begin
  Result := 'Item_' + idx.ToString;
end;

var c: IStringContainer;
begin
  c := TContainerImpl.Create;
  WriteLn(c[0]);
  WriteLn(c[1]);
end.
"#,
    );
    assert_eq!(out, vec!["Item_0", "Item_1"]);
}

#[test]
fn test_interface_multidimensional_indexer() {
    let out = run_pascal(
        r#"
program Test;
type IMatrix = interface
  ['{44444444-5555-6666-7777-888888888888}']
  function GetElem(row, col: Integer): Integer;
  property Cells[row, col: Integer]: Integer read GetElem; default;
end;

type TMatrixImpl = class(TInterfacedObject, IMatrix)
  public function GetElem(row, col: Integer): Integer;
end;
function TMatrixImpl.GetElem(row, col: Integer): Integer;
begin
  Result := row * 10 + col;
end;

var m: IMatrix;
begin
  m := TMatrixImpl.Create;
  WriteLn(m[2, 3]);
end.
"#,
    );
    assert_eq!(out, vec!["23"]);
}

#[test]
fn test_interface_write_only_property() {
    let out = run_pascal(
        r#"
program Test;
type IWriteOnlyLogger = interface
  ['{55555555-6666-7777-8888-999999999999}']
  procedure SetLogMsg(const msg: String);
  property LogMsg: String write SetLogMsg;
end;

type TLoggerImpl = class(TInterfacedObject, IWriteOnlyLogger)
  public procedure SetLogMsg(const msg: String);
end;
procedure TLoggerImpl.SetLogMsg(const msg: String);
begin
  WriteLn('Logged:' + msg);
end;

var l: IWriteOnlyLogger;
begin
  l := TLoggerImpl.Create;
  l.LogMsg := 'InterfaceWriteOnly';
end.
"#,
    );
    assert_eq!(out, vec!["Logged:InterfaceWriteOnly"]);
}

#[test]
fn test_interface_property_string_type() {
    let out = run_pascal(
        r#"
program Test;
type ITitleHolder = interface
  ['{66666666-7777-8888-9999-000000000000}']
  function GetTitle: String;
  procedure SetTitle(const t: String);
  property Title: String read GetTitle write SetTitle;
end;

type TTitleImpl = class(TInterfacedObject, ITitleHolder)
  private FTitle: String;
  public
    function GetTitle: String;
    procedure SetTitle(const t: String);
end;
function TTitleImpl.GetTitle: String; begin Result := FTitle; end;
procedure TTitleImpl.SetTitle(const t: String); begin FTitle := t; end;

var h: ITitleHolder;
begin
  h := TTitleImpl.Create;
  h.Title := 'PascalInterfaceTitle';
  WriteLn(h.Title);
end.
"#,
    );
    assert_eq!(out, vec!["PascalInterfaceTitle"]);
}

#[test]
fn test_interface_property_boolean_type() {
    let out = run_pascal(
        r#"
program Test;
type IFlagHolder = interface
  ['{77777777-8888-9999-0000-111111111111}']
  function GetFlag: Boolean;
  procedure SetFlag(v: Boolean);
  property Active: Boolean read GetFlag write SetFlag;
end;

type TFlagImpl = class(TInterfacedObject, IFlagHolder)
  private FFlag: Boolean;
  public
    function GetFlag: Boolean; procedure SetFlag(v: Boolean);
end;
function TFlagImpl.GetFlag: Boolean; begin Result := FFlag; end;
procedure TFlagImpl.SetFlag(v: Boolean); begin FFlag := v; end;

var f: IFlagHolder;
begin
  f := TFlagImpl.Create;
  f.Active := True;
  WriteLn(f.Active);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interface_property_enum_type() {
    let out = run_pascal(
        r#"
program Test;
type TState = (stStop, stRun);
type IStateHolder = interface
  ['{88888888-9999-0000-1111-222222222222}']
  function GetState: TState;
  procedure SetState(s: TState);
  property State: TState read GetState write SetState;
end;

type TStateImpl = class(TInterfacedObject, IStateHolder)
  private FState: TState;
  public function GetState: TState; procedure SetState(s: TState);
end;
function TStateImpl.GetState: TState; begin Result := FState; end;
procedure TStateImpl.SetState(s: TState); begin FState := s; end;

var s: IStateHolder;
begin
  s := TStateImpl.Create;
  s.State := stRun;
  WriteLn(Ord(s.State));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_interface_property_indexer_read_write() {
    let out = run_pascal(
        r#"
program Test;
type IMutableArray = interface
  ['{99999999-0000-1111-2222-333333333333}']
  function GetVal(idx: Integer): Integer;
  procedure SetVal(idx: Integer; v: Integer);
  property Values[idx: Integer]: Integer read GetVal write SetVal; default;
end;

type TArrImpl = class(TInterfacedObject, IMutableArray)
  private FData: array[0..2] of Integer;
  public
    function GetVal(idx: Integer): Integer;
    procedure SetVal(idx: Integer; v: Integer);
end;
function TArrImpl.GetVal(idx: Integer): Integer; begin Result := FData[idx]; end;
procedure TArrImpl.SetVal(idx: Integer; v: Integer); begin FData[idx] := v; end;

var arr: IMutableArray;
begin
  arr := TArrImpl.Create;
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  WriteLn(arr[0] + arr[1] + arr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_interface_property_record_type() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type IPointHolder = interface
  ['{00000000-1111-2222-3333-444444444444}']
  function GetPos: TPoint;
  procedure SetPos(const p: TPoint);
  property Position: TPoint read GetPos write SetPos;
end;

type TPointImpl = class(TInterfacedObject, IPointHolder)
  private FPos: TPoint;
  public function GetPos: TPoint; procedure SetPos(const p: TPoint);
end;
function TPointImpl.GetPos: TPoint; begin Result := FPos; end;
procedure TPointImpl.SetPos(const p: TPoint); begin FPos := p; end;

var h: IPointHolder; pt: TPoint;
begin
  pt.X := 15; pt.Y := 25;
  h := TPointImpl.Create;
  h.Position := pt;
  WriteLn(h.Position.X.ToString + ',' + h.Position.Y.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["15,25"]);
}

#[test]
fn test_interface_property_inheritance() {
    let out = run_pascal(
        r#"
program Test;
type IBase = interface
  ['{10101010-1010-1010-1010-101010101010}']
  function GetID: Integer;
  property ID: Integer read GetID;
end;
type ISub = interface(IBase)
  ['{20202020-2020-2020-2020-202020202020}']
  function GetName: String;
  property Name: String read GetName;
end;

type TSubImpl = class(TInterfacedObject, ISub)
  public function GetID: Integer; function GetName: String;
end;
function TSubImpl.GetID: Integer; begin Result := 1; end;
function TSubImpl.GetName: String; begin Result := 'SubName'; end;

var s: ISub;
begin
  s := TSubImpl.Create;
  WriteLn(s.ID.ToString + ':' + s.Name);
end.
"#,
    );
    assert_eq!(out, vec!["1:SubName"]);
}

#[test]
fn test_interface_property_delegation_implements() {
    let out = run_pascal(
        r#"
program Test;
type IPropIntf = interface
  ['{30303030-3030-3030-3030-303030303030}']
  function GetVal: Integer;
  property Val: Integer read GetVal;
end;

type TPropInner = class(TInterfacedObject, IPropIntf)
  public function GetVal: Integer;
end;
function TPropInner.GetVal: Integer; begin Result := 999; end;

type TPropOuter = class(TInterfacedObject, IPropIntf)
  private FInner: IPropIntf;
  public
    constructor Create;
    property Inner: IPropIntf read FInner implements IPropIntf;
end;
constructor TPropOuter.Create; begin FInner := TPropInner.Create; end;

var obj: IPropIntf;
begin
  obj := TPropOuter.Create;
  WriteLn(obj.Val);
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_interface_property_string_indexer() {
    let out = run_pascal(
        r#"
program Test;
type IDictionary = interface
  ['{40404040-4040-4040-4040-404040404040}']
  function GetVal(const key: String): String;
  property Values[const key: String]: String read GetVal; default;
end;

type TDictImpl = class(TInterfacedObject, IDictionary)
  public function GetVal(const key: String): String;
end;
function TDictImpl.GetVal(const key: String): String;
begin
  Result := 'ValFor_' + key;
end;

var dict: IDictionary;
begin
  dict := TDictImpl.Create;
  WriteLn(dict['user']);
end.
"#,
    );
    assert_eq!(out, vec!["ValFor_user"]);
}

#[test]
fn test_interface_property_double_type() {
    let out = run_pascal(
        r#"
program Test;
type IScale = interface
  ['{50505050-5050-5050-5050-505050505050}']
  function GetFactor: Double;
  procedure SetFactor(v: Double);
  property Factor: Double read GetFactor write SetFactor;
end;

type TScaleImpl = class(TInterfacedObject, IScale)
  private FFactor: Double;
  public function GetFactor: Double; procedure SetFactor(v: Double);
end;
function TScaleImpl.GetFactor: Double; begin Result := FFactor; end;
procedure TScaleImpl.SetFactor(v: Double); begin FFactor := v; end;

var s: IScale;
begin
  s := TScaleImpl.Create;
  s.Factor := 2.5;
  WriteLn(s.Factor);
end.
"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn test_interface_property_interface_type() {
    let out = run_pascal(
        r#"
program Test;
type IChildIntf = interface
  ['{60606060-6060-6060-6060-606060606060}']
  procedure Action;
end;
type IParentIntf = interface
  ['{70707070-7070-7070-7070-707070707070}']
  function GetChild: IChildIntf;
  property Child: IChildIntf read GetChild;
end;

type TChildImpl = class(TInterfacedObject, IChildIntf)
  public procedure Action;
end;
procedure TChildImpl.Action; begin WriteLn('ChildActionExecuted'); end;

type TParentImpl = class(TInterfacedObject, IParentIntf)
  private FChild: IChildIntf;
  public
    constructor Create;
    function GetChild: IChildIntf;
end;
constructor TParentImpl.Create; begin FChild := TChildImpl.Create; end;
function TParentImpl.GetChild: IChildIntf; begin Result := FChild; end;

var p: IParentIntf;
begin
  p := TParentImpl.Create;
  p.Child.Action;
end.
"#,
    );
    assert_eq!(out, vec!["ChildActionExecuted"]);
}

#[test]
fn test_interface_property_mutating_value_via_method() {
    let out = run_pascal(
        r#"
program Test;
type ICounter = interface
  ['{80808080-8080-8080-8080-808080808080}']
  function GetCount: Integer;
  procedure Increment;
  property Count: Integer read GetCount;
end;

type TCounterImpl = class(TInterfacedObject, ICounter)
  private FCount: Integer;
  public
    function GetCount: Integer; procedure Increment;
end;
function TCounterImpl.GetCount: Integer; begin Result := FCount; end;
procedure TCounterImpl.Increment; begin Inc(FCount); end;

var c: ICounter;
begin
  c := TCounterImpl.Create;
  c.Increment;
  c.Increment;
  WriteLn(c.Count);
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_interface_property_guid_access() {
    let out = run_pascal(
        r#"
program Test;
type IGuidHolder = interface
  ['{90909090-9090-9090-9090-909090909090}']
  function GetGuid: TGUID;
  property Guid: TGUID read GetGuid;
end;

type TGuidImpl = class(TInterfacedObject, IGuidHolder)
  public function GetGuid: TGUID;
end;
function TGuidImpl.GetGuid: TGUID;
begin
  Result := StringToGUID('{11111111-1111-1111-1111-111111111111}');
end;

var g: IGuidHolder;
begin
  g := TGuidImpl.Create;
  WriteLn(GUIDToString(g.Guid));
end.
"#,
    );
    assert_eq!(out, vec!["{11111111-1111-1111-1111-111111111111}"]);
}

#[test]
fn test_interface_property_array_count_indexer() {
    let out = run_pascal(
        r#"
program Test;
type ICollection = interface
  ['{A0A0A0A0-A0A0-A0A0-A0A0-A0A0A0A0A0A0}']
  function GetCount: Integer;
  function GetItem(idx: Integer): Integer;
  property Count: Integer read GetCount;
  property Items[idx: Integer]: Integer read GetItem; default;
end;

type TCollImpl = class(TInterfacedObject, ICollection)
  public function GetCount: Integer; function GetItem(idx: Integer): Integer;
end;
function TCollImpl.GetCount: Integer; begin Result := 2; end;
function TCollImpl.GetItem(idx: Integer): Integer; begin Result := (idx + 1) * 10; end;

var coll: ICollection; i: Integer;
begin
  coll := TCollImpl.Create;
  for i := 0 to coll.Count - 1 do
    WriteLn(coll[i]);
end.
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_interface_property_tbytes_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type IBufferHolder = interface
  ['{B0B0B0B0-B0B0-B0B0-B0B0-B0B0B0B0B0B0}']
  function GetBuffer: TBytes;
  property Buffer: TBytes read GetBuffer;
end;

type TBufferImpl = class(TInterfacedObject, IBufferHolder)
  public function GetBuffer: TBytes;
end;
function TBufferImpl.GetBuffer: TBytes;
begin
  SetLength(Result, 2);
  Result[0] := 65; Result[1] := 66;
end;

var b: IBufferHolder;
begin
  b := TBufferImpl.Create;
  WriteLn(Chr(b.Buffer[0]) + Chr(b.Buffer[1]));
end.
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn test_interface_property_tdatetime_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ITimeStamped = interface
  ['{C0C0C0C0-C0C0-C0C0-C0C0-C0C0C0C0C0C0}']
  function GetTimeStamp: TDateTime;
  property TimeStamp: TDateTime read GetTimeStamp;
end;

type TTimeImpl = class(TInterfacedObject, ITimeStamped)
  public function GetTimeStamp: TDateTime;
end;
function TTimeImpl.GetTimeStamp: TDateTime;
begin
  Result := EncodeDate(2026, 12, 25);
end;

var t: ITimeStamped;
begin
  t := TTimeImpl.Create;
  WriteLn(FormatDateTime('yyyy-mm-dd', t.TimeStamp));
end.
"#,
    );
    assert_eq!(out, vec!["2026-12-25"]);
}
