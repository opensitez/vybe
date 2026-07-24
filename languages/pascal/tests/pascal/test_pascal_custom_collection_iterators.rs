use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 46: Custom Collection Iterators & Enumerator Pattern
// ═══════════════════════════════════════════════════════════

#[test]
fn test_custom_range_enumerator_basic() {
    let out = run_pascal(r#"
program Test;
type TRange = record
  FStart, FEnd: Integer;
end;
type TRangeEnum = record
  private FCurr, FEnd: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TRangeEnum.MoveNext: Boolean;
begin
  Inc(FCurr);
  Result := FCurr <= FEnd;
end;
function TRangeEnum.GetCurrent: Integer; begin Result := FCurr; end;
function GetRangeEnum(r: TRange): TRangeEnum;
begin
  Result.FCurr := r.FStart - 1;
  Result.FEnd := r.FEnd;
end;

operator Enumerator(r: TRange): TRangeEnum;
begin
  Result := GetRangeEnum(r);
end;

var rng: TRange; i: Integer;
begin
  rng.FStart := 1; rng.FEnd := 3;
  for i in rng do
    WriteLn(i);
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_custom_class_getenumerator_method() {
    let out = run_pascal(r#"
program Test;
type TCustomListEnum = class
  private FItems: array[0..2] of String; FIndex: Integer;
  public constructor Create;
  public function MoveNext: Boolean;
  public function GetCurrent: String; property Current: String read GetCurrent;
end;
type TCustomList = class
  public function GetEnumerator: TCustomListEnum;
end;

constructor TCustomListEnum.Create;
begin
  FItems[0] := 'Alpha'; FItems[1] := 'Beta'; FItems[2] := 'Gamma';
  FIndex := -1;
end;
function TCustomListEnum.MoveNext: Boolean; begin Inc(FIndex); Result := FIndex <= 2; end;
function TCustomListEnum.GetCurrent: String; begin Result := FItems[FIndex]; end;
function TCustomList.GetEnumerator: TCustomListEnum; begin Result := TCustomListEnum.Create; end;

var list: TCustomList; s: String;
begin
  list := TCustomList.Create;
  for s in list do
    WriteLn(s);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["Alpha", "Beta", "Gamma"]);
}

#[test]
fn test_custom_reverse_enumerator() {
    let out = run_pascal(r#"
program Test;
type TArrayWrapper = record
  Data: array[0..2] of Integer;
end;
type TReverseEnum = record
  private FData: array[0..2] of Integer; FIndex: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TReverseEnum.MoveNext: Boolean; begin Dec(FIndex); Result := FIndex >= 0; end;
function TReverseEnum.GetCurrent: Integer; begin Result := FData[FIndex]; end;

operator Enumerator(w: TArrayWrapper): TReverseEnum;
begin
  Result.FData := w.Data;
  Result.FIndex := 3;
end;

var wrapper: TArrayWrapper; val: Integer;
begin
  wrapper.Data[0] := 10; wrapper.Data[1] := 20; wrapper.Data[2] := 30;
  for val in wrapper do
    WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn test_custom_filter_enumerator_even_numbers() {
    let out = run_pascal(r#"
program Test;
type TEvenFilter = record
  MaxVal: Integer;
end;
type TEvenEnum = record
  private FCurr, FMax: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TEvenEnum.MoveNext: Boolean;
begin
  Inc(FCurr, 2);
  Result := FCurr <= FMax;
end;
function TEvenEnum.GetCurrent: Integer; begin Result := FCurr; end;

operator Enumerator(f: TEvenFilter): TEvenEnum;
begin
  Result.FCurr := 0;
  Result.FMax := f.MaxVal;
end;

var filter: TEvenFilter; val: Integer;
begin
  filter.MaxVal := 6;
  for val in filter do
    WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["2", "4", "6"]);
}

#[test]
fn test_linked_list_enumerator() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record
       Val: Integer;
       Next: PNode;
     end;

type TListWrapper = record
  Head: PNode;
end;

type TNodeEnum = record
  private FCurr: PNode;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;

function TNodeEnum.MoveNext: Boolean;
begin
  if FCurr <> nil then FCurr := FCurr^.Next;
  Result := FCurr <> nil;
end;
function TNodeEnum.GetCurrent: Integer; begin Result := FCurr^.Val; end;

operator Enumerator(w: TListWrapper): TNodeEnum;
var dummyHeader: TNode;
begin
  dummyHeader.Val := 0; dummyHeader.Next := w.Head;
  Result.FCurr := @dummyHeader;
end;

var n1, n2: PNode; wrap: TListWrapper; v: Integer;
begin
  New(n1); New(n2);
  n1^.Val := 100; n1^.Next := n2;
  n2^.Val := 200; n2^.Next := nil;
  wrap.Head := n1;
  for v in wrap do
    WriteLn(v);
  Dispose(n1); Dispose(n2);
end.
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_custom_generic_container_enumerator() {
    let out = run_pascal(r#"
program Test;
type TGenBox<T> = class
  public Item1, Item2: T;
  constructor Create(v1, v2: T);
end;
type TGenBoxEnum<T> = record
  private FVal1, FVal2: T; FStep: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: T; property Current: T read GetCurrent;
end;

constructor TGenBox<T>.Create(v1, v2: T); begin Item1 := v1; Item2 := v2; end;
function TGenBoxEnum<T>.MoveNext: Boolean; begin Inc(FStep); Result := FStep <= 2; end;
function TGenBoxEnum<T>.GetCurrent: T; begin if FStep = 1 then Result := FVal1 else Result := FVal2; end;

function GetBoxEnum<T>(box: TGenBox<T>): TGenBoxEnum<T>;
begin
  Result.FVal1 := box.Item1; Result.FVal2 := box.Item2; Result.FStep := 0;
end;

operator Enumerator<T>(box: TGenBox<T>): TGenBoxEnum<T>;
begin
  Result := GetBoxEnum<T>(box);
end;

var b: TGenBox<String>; s: String;
begin
  b := TGenBox<String>.Create('BoxA', 'BoxB');
  for s in b do
    WriteLn(s);
  b.Free;
end.
"#);
    assert_eq!(out, vec!["BoxA", "BoxB"]);
}

#[test]
fn test_custom_enumerator_with_step_size() {
    let out = run_pascal(r#"
program Test;
type TStepRange = record
  Start, Stop, Step: Integer;
end;
type TStepEnum = record
  private FCurr, FStop, FStep: Integer; FFirst: Boolean;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;

function TStepEnum.MoveNext: Boolean;
begin
  if FFirst then FFirst := False
  else FCurr := FCurr + FStep;
  Result := FCurr <= FStop;
end;
function TStepEnum.GetCurrent: Integer; begin Result := FCurr; end;

operator Enumerator(sr: TStepRange): TStepEnum;
begin
  Result.FCurr := sr.Start; Result.FStop := sr.Stop; Result.FStep := sr.Step; Result.FFirst := True;
end;

var sr: TStepRange; val: Integer;
begin
  sr.Start := 10; sr.Stop := 30; sr.Step := 10;
  for val in sr do
    WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_custom_enumerator_yields_index_value_pair() {
    let out = run_pascal(r#"
program Test;
type TPair = record Index: Integer; Val: String; end;
type TIndexedArr = record
  Items: array[0..1] of String;
end;
type TIndexedEnum = record
  private FItems: array[0..1] of String; FIdx: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: TPair; property Current: TPair read GetCurrent;
end;

function TIndexedEnum.MoveNext: Boolean; begin Inc(FIdx); Result := FIdx <= 1; end;
function TIndexedEnum.GetCurrent: TPair; begin Result.Index := FIdx; Result.Val := FItems[FIdx]; end;

operator Enumerator(ia: TIndexedArr): TIndexedEnum;
begin
  Result.FItems := ia.Items; Result.FIdx := -1;
end;

var arr: TIndexedArr; p: TPair;
begin
  arr.Items[0] := 'Zero'; arr.Items[1] := 'One';
  for p in arr do
    WriteLn(p.Index.ToString + '=' + p.Val);
end.
"#);
    assert_eq!(out, vec!["0=Zero", "1=One"]);
}

#[test]
fn test_custom_string_token_enumerator() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TTokenWrapper = record
  Text: String; Delim: Char;
end;
type TTokenEnum = record
  private FRest: String; FCurr: String; FDelim: Char;
  public function MoveNext: Boolean;
  public function GetCurrent: String; property Current: String read GetCurrent;
end;

function TTokenEnum.MoveNext: Boolean;
var p: Integer;
begin
  if FRest = '' then Exit(False);
  p := Pos(FDelim, FRest);
  if p > 0 then
  begin
    FCurr := Copy(FRest, 1, p - 1);
    Delete(FRest, 1, p);
  end else
  begin
    FCurr := FRest;
    FRest := '';
  end;
  Result := True;
end;
function TTokenEnum.GetCurrent: String; begin Result := FCurr; end;

operator Enumerator(tw: TTokenWrapper): TTokenEnum;
begin
  Result.FRest := tw.Text; Result.FDelim := tw.Delim;
end;

var tw: TTokenWrapper; tok: String;
begin
  tw.Text := 'red,green,blue'; tw.Delim := ',';
  for tok in tw do
    WriteLn(tok);
end.
"#);
    assert_eq!(out, vec!["red", "green", "blue"]);
}

#[test]
fn test_custom_enumerator_empty_container() {
    let out = run_pascal(r#"
program Test;
type TEmptyContainer = record end;
type TEmptyEnum = record
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TEmptyEnum.MoveNext: Boolean; begin Result := False; end;
function TEmptyEnum.GetCurrent: Integer; begin Result := 0; end;

operator Enumerator(c: TEmptyContainer): TEmptyEnum;
begin end;

var ec: TEmptyContainer; val, count: Integer;
begin
  count := 0;
  for val in ec do Inc(count);
  WriteLn(count);
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_custom_enumerator_single_item() {
    let out = run_pascal(r#"
program Test;
type TSingleContainer = record Val: Integer; end;
type TSingleEnum = record
  private FVal: Integer; FDone: Boolean;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TSingleEnum.MoveNext: Boolean;
begin
  if not FDone then begin FDone := True; Result := True; end else Result := False;
end;
function TSingleEnum.GetCurrent: Integer; begin Result := FVal; end;

operator Enumerator(sc: TSingleContainer): TSingleEnum;
begin
  Result.FVal := sc.Val; Result.FDone := False;
end;

var sc: TSingleContainer; v: Integer;
begin
  sc.Val := 999;
  for v in sc do
    WriteLn(v);
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_custom_enumerator_record_fields() {
    let out = run_pascal(r#"
program Test;
type TPointRec = record X, Y, Z: Integer; end;
type TPointEnum = record
  private FRec: TPointRec; FIdx: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TPointEnum.MoveNext: Boolean; begin Inc(FIdx); Result := FIdx <= 3; end;
function TPointEnum.GetCurrent: Integer;
begin
  case FIdx of
    1: Result := FRec.X;
    2: Result := FRec.Y;
    3: Result := FRec.Z;
  else Result := 0;
  end;
end;

operator Enumerator(pt: TPointRec): TPointEnum;
begin
  Result.FRec := pt; Result.FIdx := 0;
end;

var pt: TPointRec; coord: Integer;
begin
  pt.X := 10; pt.Y := 20; pt.Z := 30;
  for coord in pt do
    WriteLn(coord);
end.
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_custom_enumerator_matrix_row_traversal() {
    let out = run_pascal(r#"
program Test;
type TMatrix2x2 = record
  Data: array[0..1, 0..1] of Integer;
end;
type TMatEnum = record
  private FData: array[0..1, 0..1] of Integer; FIdx: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TMatEnum.MoveNext: Boolean; begin Inc(FIdx); Result := FIdx <= 3; end;
function TMatEnum.GetCurrent: Integer;
begin
  Result := FData[FIdx div 2, FIdx mod 2];
end;

operator Enumerator(m: TMatrix2x2): TMatEnum;
begin
  Result.FData := m.Data; Result.FIdx := -1;
end;

var m: TMatrix2x2; v, sum: Integer;
begin
  m.Data[0,0] := 1; m.Data[0,1] := 2;
  m.Data[1,0] := 3; m.Data[1,1] := 4;
  sum := 0;
  for v in m do sum := sum + v;
  WriteLn(sum);
end.
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_custom_enumerator_boolean_flags() {
    let out = run_pascal(r#"
program Test;
type TFlagContainer = record
  Flags: array[0..2] of Boolean;
end;
type TFlagEnum = record
  private FFlags: array[0..2] of Boolean; FIdx: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Boolean; property Current: Boolean read GetCurrent;
end;
function TFlagEnum.MoveNext: Boolean; begin Inc(FIdx); Result := FIdx <= 2; end;
function TFlagEnum.GetCurrent: Boolean; begin Result := FFlags[FIdx]; end;

operator Enumerator(fc: TFlagContainer): TFlagEnum;
begin
  Result.FFlags := fc.Flags; Result.FIdx := -1;
end;

var fc: TFlagContainer; b: Boolean;
begin
  fc.Flags[0] := True; fc.Flags[1] := False; fc.Flags[2] := True;
  for b in fc do
    WriteLn(b);
end.
"#);
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_custom_enumerator_break_loop_support() {
    let out = run_pascal(r#"
program Test;
type TRange = record StartVal, StopVal: Integer; end;
type TRangeEnum = record
  private FCurr, FStop: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TRangeEnum.MoveNext: Boolean; begin Inc(FCurr); Result := FCurr <= FStop; end;
function TRangeEnum.GetCurrent: Integer; begin Result := FCurr; end;

operator Enumerator(r: TRange): TRangeEnum;
begin
  Result.FCurr := r.StartVal - 1; Result.FStop := r.StopVal;
end;

var r: TRange; i: Integer;
begin
  r.StartVal := 1; r.StopVal := 10;
  for i in r do
  begin
    if i = 4 then Break;
    WriteLn(i);
  end;
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_custom_enumerator_continue_loop_support() {
    let out = run_pascal(r#"
program Test;
type TRange = record StartVal, StopVal: Integer; end;
type TRangeEnum = record
  private FCurr, FStop: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TRangeEnum.MoveNext: Boolean; begin Inc(FCurr); Result := FCurr <= FStop; end;
function TRangeEnum.GetCurrent: Integer; begin Result := FCurr; end;

operator Enumerator(r: TRange): TRangeEnum;
begin
  Result.FCurr := r.StartVal - 1; Result.FStop := r.StopVal;
end;

var r: TRange; i: Integer;
begin
  r.StartVal := 1; r.StopVal := 5;
  for i in r do
  begin
    if (i mod 2) = 0 then Continue;
    WriteLn(i);
  end;
end.
"#);
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn test_custom_enumerator_nested_loops() {
    let out = run_pascal(r#"
program Test;
type TRange = record StartVal, StopVal: Integer; end;
type TRangeEnum = record
  private FCurr, FStop: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function TRangeEnum.MoveNext: Boolean; begin Inc(FCurr); Result := FCurr <= FStop; end;
function TRangeEnum.GetCurrent: Integer; begin Result := FCurr; end;

operator Enumerator(r: TRange): TRangeEnum;
begin
  Result.FCurr := r.StartVal - 1; Result.FStop := r.StopVal;
end;

var r1, r2: TRange; i, j: Integer;
begin
  r1.StartVal := 1; r1.StopVal := 2;
  r2.StartVal := 10; r2.StopVal := 11;
  for i in r1 do
    for j in r2 do
      WriteLn(i.ToString + ':' + j.ToString);
end.
"#);
    assert_eq!(out, vec!["1:10", "1:11", "2:10", "2:11"]);
}

#[test]
fn test_custom_enumerator_enum_type() {
    let out = run_pascal(r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
type TColorEnum = record
  private FCurr: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: TColor; property Current: TColor read GetCurrent;
end;
function TColorEnum.MoveNext: Boolean; begin Inc(FCurr); Result := FCurr <= Ord(High(TColor)); end;
function TColorEnum.GetCurrent: TColor; begin Result := TColor(FCurr); end;

type TColorRange = record end;
operator Enumerator(cr: TColorRange): TColorEnum;
begin
  Result.FCurr := -1;
end;

var cr: TColorRange; col: TColor;
begin
  for col in cr do
    WriteLn(Ord(col));
end.
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn test_custom_enumerator_real_elements() {
    let out = run_pascal(r#"
program Test;
type TRealList = record
  R1, R2: Real;
end;
type TRealEnum = record
  private FVal1, FVal2: Real; FStep: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Real; property Current: Real read GetCurrent;
end;
function TRealEnum.MoveNext: Boolean; begin Inc(FStep); Result := FStep <= 2; end;
function TRealEnum.GetCurrent: Real; begin if FStep = 1 then Result := FVal1 else Result := FVal2; end;

operator Enumerator(rl: TRealList): TRealEnum;
begin
  Result.FVal1 := rl.R1; Result.FVal2 := rl.R2; Result.FStep := 0;
end;

var rl: TRealList; r: Real;
begin
  rl.R1 := 2.5; rl.R2 := 7.5;
  for r in rl do
    WriteLn(r);
end.
"#);
    assert_eq!(out, vec!["2.5", "7.5"]);
}

#[test]
fn test_custom_class_helper_enumerator() {
    let out = run_pascal(r#"
program Test;
type THelperEnum = record
  private FCurr, FMax: Integer;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;
function THelperEnum.MoveNext: Boolean; begin Inc(FCurr); Result := FCurr <= FMax; end;
function THelperEnum.GetCurrent: Integer; begin Result := FCurr; end;

type TIntegerHelper = record helper for Integer
  public function GetEnumerator: THelperEnum;
end;
function TIntegerHelper.GetEnumerator: THelperEnum;
begin
  Result.FCurr := 0; Result.FMax := Self;
end;

var num, i: Integer;
begin
  num := 3;
  for i in num do
    WriteLn(i);
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}
