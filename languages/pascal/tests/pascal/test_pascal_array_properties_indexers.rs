use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 13: Array Properties & Indexers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_array_property_default_bracket_indexer() {
    let out = run_pascal(
        r#"
program Test;
type TIntList = class
  private FData: array[0..4] of Integer;
  private function GetItem(index: Integer): Integer;
  private procedure SetItem(index, val: Integer);
  public property Items[index: Integer]: Integer read GetItem write SetItem; default;
end;
function TIntList.GetItem(index: Integer): Integer; begin Result := FData[index]; end;
procedure TIntList.SetItem(index, val: Integer); begin FData[index] := val; end;
var list: TIntList;
begin
  list := TIntList.Create;
  list[0] := 10;
  list[1] := 20;
  WriteLn(list[0]);
  WriteLn(list[1]);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_array_property_explicit_name_access() {
    let out = run_pascal(
        r#"
program Test;
type TStrList = class
  private FItems: array[0..2] of String;
  private function GetItem(i: Integer): String;
  private procedure SetItem(i: Integer; s: String);
  public property Strings[i: Integer]: String read GetItem write SetItem;
end;
function TStrList.GetItem(i: Integer): String; begin Result := FItems[i]; end;
procedure TStrList.SetItem(i: Integer; s: String); begin FItems[i] := s; end;
var list: TStrList;
begin
  list := TStrList.Create;
  list.Strings[0] := 'First';
  WriteLn(list.Strings[0]);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["First"]);
}

#[test]
fn test_multidimensional_array_property() {
    let out = run_pascal(
        r#"
program Test;
type TGrid = class
  private FCells: array[0..2, 0..2] of Integer;
  private function GetCell(row, col: Integer): Integer;
  private procedure SetCell(row, col, val: Integer);
  public property Cells[row, col: Integer]: Integer read GetCell write SetCell; default;
end;
function TGrid.GetCell(row, col: Integer): Integer; begin Result := FCells[row, col]; end;
procedure TGrid.SetCell(row, col, val: Integer); begin FCells[row, col] := val; end;
var g: TGrid;
begin
  g := TGrid.Create;
  g[1, 2] := 99;
  WriteLn(g[1, 2]);
  g.Free;
end.
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_string_indexed_array_property() {
    let out = run_pascal(
        r#"
program Test;
type TMapHolder = class
  private FKeys: array[0..2] of String;
  private FValues: array[0..2] of String;
  private function GetVal(key: String): String;
  private procedure SetVal(key, val: String);
  public constructor Create;
  public property Values[key: String]: String read GetVal write SetVal; default;
end;
constructor TMapHolder.Create;
begin
  FKeys[0] := 'host'; FKeys[1] := 'port';
end;
function TMapHolder.GetVal(key: String): String;
begin
  if key = 'host' then Result := FValues[0]
  else if key = 'port' then Result := FValues[1]
  else Result := '';
end;
procedure TMapHolder.SetVal(key, val: String);
begin
  if key = 'host' then FValues[0] := val
  else if key = 'port' then FValues[1] := val;
end;
var map: TMapHolder;
begin
  map := TMapHolder.Create;
  map['host'] := 'localhost';
  map['port'] := '8080';
  WriteLn(map['host'] + ':' + map['port']);
  map.Free;
end.
"#,
    );
    assert_eq!(out, vec!["localhost:8080"]);
}

#[test]
fn test_enum_indexed_array_property() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
type TPalette = class
  private FColors: array[TColor] of String;
  private function GetHex(c: TColor): String;
  private procedure SetHex(c: TColor; hex: String);
  public property ColorHex[c: TColor]: String read GetHex write SetHex; default;
end;
function TPalette.GetHex(c: TColor): String; begin Result := FColors[c]; end;
procedure TPalette.SetHex(c: TColor; hex: String); begin FColors[c] := hex; end;
var p: TPalette;
begin
  p := TPalette.Create;
  p[cRed] := '#FF0000';
  p[cGreen] := '#00FF00';
  WriteLn(p[cRed]);
  WriteLn(p[cGreen]);
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["#FF0000", "#00FF00"]);
}

#[test]
fn test_read_only_array_property() {
    let out = run_pascal(
        r#"
program Test;
type TFixedList = class
  private FNumbers: array[0..2] of Integer;
  private function GetNumber(i: Integer): Integer;
  public constructor Create;
  public property Numbers[i: Integer]: Integer read GetNumber; default;
end;
constructor TFixedList.Create;
begin
  FNumbers[0] := 100; FNumbers[1] := 200; FNumbers[2] := 300;
end;
function TFixedList.GetNumber(i: Integer): Integer; begin Result := FNumbers[i]; end;
var list: TFixedList;
begin
  list := TFixedList.Create;
  WriteLn(list[1]);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["200"]);
}

#[test]
fn test_write_only_array_property() {
    let out = run_pascal(
        r#"
program Test;
type TBufferWriter = class
  private FBuf: array[0..2] of Integer;
  private procedure SetBuf(i: Integer; val: Integer);
  public property Buffer[i: Integer]: Integer write SetBuf;
  public function GetSum: Integer;
end;
procedure TBufferWriter.SetBuf(i: Integer; val: Integer); begin FBuf[i] := val; end;
function TBufferWriter.GetSum: Integer; begin Result := FBuf[0] + FBuf[1] + FBuf[2]; end;
var bw: TBufferWriter;
begin
  bw := TBufferWriter.Create;
  bw.Buffer[0] := 10;
  bw.Buffer[1] := 20;
  bw.Buffer[2] := 30;
  WriteLn(bw.GetSum);
  bw.Free;
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_array_property_bounds_checking_setter() {
    let out = run_pascal(
        r#"
program Test;
type TBoundedArray = class
  private FItems: array[0..2] of Integer;
  private procedure SetItem(i: Integer; val: Integer);
  private function GetItem(i: Integer): Integer;
  public property Items[i: Integer]: Integer read GetItem write SetItem; default;
end;
procedure TBoundedArray.SetItem(i: Integer; val: Integer);
begin
  if (i >= 0) and (i <= 2) then FItems[i] := val;
end;
function TBoundedArray.GetItem(i: Integer): Integer;
begin
  if (i >= 0) and (i <= 2) then Result := FItems[i] else Result := -1;
end;
var ba: TBoundedArray;
begin
  ba := TBoundedArray.Create;
  ba[0] := 42;
  ba[5] := 99;
  WriteLn(ba[0]);
  WriteLn(ba[5]);
  ba.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42", "-1"]);
}

#[test]
fn test_array_property_returning_record() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TPointArray = class
  private FPoints: array[0..1] of TPoint;
  private function GetPoint(i: Integer): TPoint;
  private procedure SetPoint(i: Integer; p: TPoint);
  public property Points[i: Integer]: TPoint read GetPoint write SetPoint; default;
end;
function TPointArray.GetPoint(i: Integer): TPoint; begin Result := FPoints[i]; end;
procedure TPointArray.SetPoint(i: Integer; p: TPoint); begin FPoints[i] := p; end;
var pa: TPointArray; pt: TPoint;
begin
  pa := TPointArray.Create;
  pt.X := 10; pt.Y := 20;
  pa[0] := pt;
  WriteLn(pa[0].X);
  WriteLn(pa[0].Y);
  pa.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_array_property_subrange_index_type() {
    let out = run_pascal(
        r#"
program Test;
type TIndex = 1..3;
type TSubrangeIndexed = class
  private FData: array[TIndex] of String;
  private function GetVal(idx: TIndex): String;
  private procedure SetVal(idx: TIndex; val: String);
  public property Values[idx: TIndex]: String read GetVal write SetVal; default;
end;
function TSubrangeIndexed.GetVal(idx: TIndex): String; begin Result := FData[idx]; end;
procedure TSubrangeIndexed.SetVal(idx: TIndex; val: String); begin FData[idx] := val; end;
var si: TSubrangeIndexed;
begin
  si := TSubrangeIndexed.Create;
  si[1] := 'One';
  si[3] := 'Three';
  WriteLn(si[1]);
  WriteLn(si[3]);
  si.Free;
end.
"#,
    );
    assert_eq!(out, vec!["One", "Three"]);
}

#[test]
fn test_array_property_virtual_accessor_methods() {
    let out = run_pascal(
        r#"
program Test;
type TBaseContainer = class
  protected function GetElem(i: Integer): String; virtual;
  public property Elem[i: Integer]: String read GetElem; default;
end;
type TDerivedContainer = class(TBaseContainer)
  protected function GetElem(i: Integer): String; override;
end;
function TBaseContainer.GetElem(i: Integer): String; begin Result := 'Base'; end;
function TDerivedContainer.GetElem(i: Integer): String; begin Result := 'Derived:' + i.ToString; end;
var c: TBaseContainer;
begin
  c := TDerivedContainer.Create;
  WriteLn(c[5]);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Derived:5"]);
}

#[test]
fn test_array_property_combined_with_count_property() {
    let out = run_pascal(
        r#"
program Test;
type TFixedStack = class
  private FData: array[0..4] of Integer; FCount: Integer;
  private function GetItem(i: Integer): Integer;
  public constructor Create; procedure Push(v: Integer);
  public property Count: Integer read FCount;
  public property Items[i: Integer]: Integer read GetItem; default;
end;
constructor TFixedStack.Create; begin FCount := 0; end;
procedure TFixedStack.Push(v: Integer); begin FData[FCount] := v; Inc(FCount); end;
function TFixedStack.GetItem(i: Integer): Integer; begin Result := FData[i]; end;
var s: TFixedStack; i: Integer;
begin
  s := TFixedStack.Create;
  s.Push(10); s.Push(20); s.Push(30);
  for i := 0 to s.Count - 1 do
    WriteLn(s[i]);
  s.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_array_property_modifying_internal_counter() {
    let out = run_pascal(
        r#"
program Test;
type TTrackedArray = class
  private FData: array[0..2] of Integer; FMutations: Integer;
  private procedure SetVal(i, val: Integer);
  private function GetVal(i: Integer): Integer;
  public constructor Create;
  public property Items[i: Integer]: Integer read GetVal write SetVal; default;
  public property Mutations: Integer read FMutations;
end;
constructor TTrackedArray.Create; begin FMutations := 0; end;
procedure TTrackedArray.SetVal(i, val: Integer); begin FData[i] := val; Inc(FMutations); end;
function TTrackedArray.GetVal(i: Integer): Integer; begin Result := FData[i]; end;
var ta: TTrackedArray;
begin
  ta := TTrackedArray.Create;
  ta[0] := 5; ta[1] := 10;
  WriteLn(ta.Mutations);
  ta.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_array_property_overloaded_index_types() {
    let out = run_pascal(
        r#"
program Test;
type TFlexContainer = class
  private FData: array[0..2] of String;
  private function GetByIdx(i: Integer): String;
  private function GetByName(name: String): String;
  public constructor Create;
  public property Items[i: Integer]: String read GetByIdx; default;
  public property Items[name: String]: String read GetByName; default;
end;
constructor TFlexContainer.Create; begin FData[0] := 'Alpha'; FData[1] := 'Beta'; end;
function TFlexContainer.GetByIdx(i: Integer): String; begin Result := FData[i]; end;
function TFlexContainer.GetByName(name: String): String;
begin
  if name = 'first' then Result := FData[0] else Result := FData[1];
end;
var fc: TFlexContainer;
begin
  fc := TFlexContainer.Create;
  WriteLn(fc[0]);
  WriteLn(fc['first']);
  fc.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Alpha", "Alpha"]);
}

#[test]
fn test_array_property_returning_boolean() {
    let out = run_pascal(
        r#"
program Test;
type TBitFlags = class
  private FFlags: array[0..7] of Boolean;
  private function GetBit(i: Integer): Boolean;
  private procedure SetBit(i: Integer; b: Boolean);
  public property Bits[i: Integer]: Boolean read GetBit write SetBit; default;
end;
function TBitFlags.GetBit(i: Integer): Boolean; begin Result := FFlags[i]; end;
procedure TBitFlags.SetBit(i: Integer; b: Boolean); begin FFlags[i] := b; end;
var bf: TBitFlags;
begin
  bf := TBitFlags.Create;
  bf[3] := True;
  WriteLn(bf[0]);
  WriteLn(bf[3]);
  bf.Free;
end.
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_array_property_in_nested_loops() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix2x2 = class
  private FData: array[0..1, 0..1] of Integer;
  private function GetV(r, c: Integer): Integer;
  private procedure SetV(r, c, v: Integer);
  public property Cells[r, c: Integer]: Integer read GetV write SetV; default;
end;
function TMatrix2x2.GetV(r, c: Integer): Integer; begin Result := FData[r, c]; end;
procedure TMatrix2x2.SetV(r, c, v: Integer); begin FData[r, c] := v; end;
var m: TMatrix2x2; r, c, sum: Integer;
begin
  m := TMatrix2x2.Create;
  m[0, 0] := 1; m[0, 1] := 2;
  m[1, 0] := 3; m[1, 1] := 4;
  sum := 0;
  for r := 0 to 1 do
    for c := 0 to 1 do
      sum := sum + m[r, c];
  WriteLn(sum);
  m.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_array_property_inherited_redeclaration() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  protected FValues: array[0..2] of Integer;
  protected function GetVal(i: Integer): Integer;
  public property Values[i: Integer]: Integer read GetVal; default;
end;
type TDerived = class(TBase)
  protected procedure SetVal(i, v: Integer);
  public property Values[i: Integer]: Integer read GetVal write SetVal; default;
end;
function TBase.GetVal(i: Integer): Integer; begin Result := FValues[i]; end;
procedure TDerived.SetVal(i, v: Integer); begin FValues[i] := v; end;
var d: TDerived;
begin
  d := TDerived.Create;
  d[1] := 77;
  WriteLn(d[1]);
  d.Free;
end.
"#,
    );
    assert_eq!(out, vec!["77"]);
}

#[test]
fn test_array_property_string_concatenation_setter() {
    let out = run_pascal(
        r#"
program Test;
type TLogBuffer = class
  private FLogs: array[0..2] of String;
  private procedure AppendLog(i: Integer; msg: String);
  private function GetLog(i: Integer): String;
  public property Logs[i: Integer]: String read GetLog write AppendLog; default;
end;
procedure TLogBuffer.AppendLog(i: Integer; msg: String);
begin
  FLogs[i] := FLogs[i] + msg;
end;
function TLogBuffer.GetLog(i: Integer): String; begin Result := FLogs[i]; end;
var lb: TLogBuffer;
begin
  lb := TLogBuffer.Create;
  lb[0] := 'Line1:';
  lb[0] := ' OK';
  WriteLn(lb[0]);
  lb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Line1: OK"]);
}

#[test]
fn test_array_property_object_instance_element() {
    let out = run_pascal(
        r#"
program Test;
type TNode = class public LabelStr: String; constructor Create(L: String); end;
type TNodeMap = class
  private FNodes: array[0..1] of TNode;
  private function GetNode(i: Integer): TNode;
  private procedure SetNode(i: Integer; n: TNode);
  public property Nodes[i: Integer]: TNode read GetNode write SetNode; default;
  public destructor Destroy; override;
end;
constructor TNode.Create(L: String); begin LabelStr := L; end;
function TNodeMap.GetNode(i: Integer): TNode; begin Result := FNodes[i]; end;
procedure TNodeMap.SetNode(i: Integer; n: TNode); begin FNodes[i] := n; end;
destructor TNodeMap.Destroy; begin FNodes[0].Free; FNodes[1].Free; inherited Destroy; end;
var nm: TNodeMap;
begin
  nm := TNodeMap.Create;
  nm[0] := TNode.Create('Start');
  nm[1] := TNode.Create('End');
  WriteLn(nm[0].LabelStr + '-' + nm[1].LabelStr);
  nm.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Start-End"]);
}

#[test]
fn test_array_property_float_values() {
    let out = run_pascal(
        r#"
program Test;
type TFloatList = class
  private FFloats: array[0..2] of Real;
  private function GetF(i: Integer): Real;
  private procedure SetF(i: Integer; v: Real);
  public property Floats[i: Integer]: Real read GetF write SetF; default;
end;
function TFloatList.GetF(i: Integer): Real; begin Result := FFloats[i]; end;
procedure TFloatList.SetF(i: Integer; v: Real); begin FFloats[i] := v; end;
var fl: TFloatList;
begin
  fl := TFloatList.Create;
  fl[0] := 1.1; fl[1] := 2.2;
  WriteLn(fl[0] + fl[1]);
  fl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3.3"]);
}
