/// Extended property patterns: indexed, default, dynamic array, class properties.
use super::helpers::run_pascal;

#[test]
fn property_dynamic_array_length_read() {
    assert_eq!(
        run_pascal(r#"program T; type TBag=class private F:array of Integer; public property Count:Integer read GetCount; function GetCount:Integer; begin Result:=Length(F); end; procedure Add(v:Integer); var n:Integer; begin n:=Length(F); SetLength(F,n+1); F[n]:=v; end; end; var b:TBag; begin b:=TBag.Create; b.Add(3); b.Add(7); WriteLn(b.Count); end."#),
        &["2"]
    );
}

#[test]
fn property_dynamic_array_indexed_write_read() {
    assert_eq!(
        run_pascal(r#"program T; type TBuf=class private F:array of string; public property Cells[i:Integer]:string read GetCell write SetCell; function GetCell(i:Integer):string; begin Result:=F[i]; end; procedure SetCell(i:Integer; const v:string); begin F[i]:=v; end; procedure Grow(n:Integer); begin SetLength(F,n); end; end; var b:TBuf; begin b:=TBuf.Create; b.Grow(2); b.Cells[0]:='a'; b.Cells[1]:='b'; WriteLn(b.Cells[1]); end."#),
        &["b"]
    );
}

#[test]
fn property_default_allows_bracket_syntax() {
    assert_eq!(
        run_pascal(r#"program T; type TMap=class private F:array[0..1] of Integer; public property Items[i:Integer]:Integer read GetItem write SetItem; default; function GetItem(i:Integer):Integer; begin Result:=F[i]; end; procedure SetItem(i:Integer; v:Integer); begin F[i]:=v; end; end; var m:TMap; begin m:=TMap.Create; m[0]:=9; m[1]:=4; WriteLn(m[1]); end."#),
        &["4"]
    );
}

#[test]
fn property_class_static_read_write() {
    assert_eq!(
        run_pascal(r#"program T; type TApp=class strict private class var FName:string; class function GetName:string; class procedure SetName(const v:string); public class property AppName:string read GetName write SetName; end; class function TApp.GetName:string; begin Result:=FName; end; class procedure TApp.SetName(const v:string); begin FName:=v; end; begin TApp.AppName:='vybe'; WriteLn(TApp.AppName); end."#),
        &["vybe"]
    );
}

#[test]
fn property_write_only_via_setter_side_effect() {
    assert_eq!(
        run_pascal(r#"program T; type TLog=class private FLast:Integer; procedure SetTarget(v:Integer); public property Target:Integer write SetTarget; function Last:Integer; begin Result:=FLast; end; end; procedure TLog.SetTarget(v:Integer); begin FLast:=v; end; var L:TLog; begin L:=TLog.Create; L.Target:=42; WriteLn(L.Last); end."#),
        &["42"]
    );
}

#[test]
fn property_readonly_exposes_field_after_ctor() {
    assert_eq!(
        run_pascal(r#"program T; type TToken=class private FKind:Integer; public constructor Create(k:Integer); property Kind:Integer read FKind; end; constructor TToken.Create(k:Integer); begin FKind:=k; end; var t:TToken; begin t:=TToken.Create(7); WriteLn(t.Kind); end."#),
        &["7"]
    );
}

#[test]
fn property_inherited_indexed_in_child() {
    assert_eq!(
        run_pascal(r#"program T; type TBase=class private F:array[0..1] of Integer; protected function GetItem(i:Integer):Integer; procedure SetItem(i:Integer; v:Integer); public property Data[i:Integer]:Integer read GetItem write SetItem; end; TChild=class(TBase); function TBase.GetItem(i:Integer):Integer; begin Result:=F[i]; end; procedure TBase.SetItem(i:Integer; v:Integer); begin F[i]:=v; end; var c:TChild; begin c:=TChild.Create; c.Data[0]:=11; WriteLn(c.Data[0]); end."#),
        &["11"]
    );
}

#[test]
fn property_getter_computes_from_two_fields() {
    assert_eq!(
        run_pascal(r#"program T; type TRect=class private FW,FH:Integer; function GetArea:Integer; public property Width:Integer read FW write FW; property Height:Integer read FH write FH; property Area:Integer read GetArea; end; function TRect.GetArea:Integer; begin Result:=FW*FH; end; var r:TRect; begin r:=TRect.Create; r.Width:=4; r.Height:=5; WriteLn(r.Area); end."#),
        &["20"]
    );
}

#[test]
fn property_setter_clamps_upper_bound() {
    assert_eq!(
        run_pascal(r#"program T; type TPercent=class private F:Integer; procedure SetP(v:Integer); public property P:Integer read F write SetP; end; procedure TPercent.SetP(v:Integer); begin if v>100 then F:=100 else F:=v; end; var p:TPercent; begin p:=TPercent.Create; p.P:=150; WriteLn(p.P); end."#),
        &["100"]
    );
}

#[test]
fn property_setter_clamps_lower_bound() {
    assert_eq!(
        run_pascal(r#"program T; type TPercent=class private F:Integer; procedure SetP(v:Integer); public property P:Integer read F write SetP; end; procedure TPercent.SetP(v:Integer); begin if v<0 then F:=0 else F:=v; end; var p:TPercent; begin p:=TPercent.Create; p.P:=-5; WriteLn(p.P); end."#),
        &["0"]
    );
}

#[test]
fn property_string_uppercase_computed() {
    assert_eq!(
        run_pascal(r#"program T; type TLabel=class private FText:string; function GetUpper:string; public property Text:string read FText write FText; property Upper:string read GetUpper; end; function TLabel.GetUpper:string; begin Result:=UpperCase(FText); end; var L:TLabel; begin L:=TLabel.Create; L.Text:='hi'; WriteLn(L.Upper); end."#),
        &["HI"]
    );
}

#[test]
fn property_bool_is_empty_on_string() {
    assert_eq!(
        run_pascal(r#"program T; type TName=class private F:string; function GetEmpty:Boolean; public property Value:string read F write F; property IsEmpty:Boolean read GetEmpty; end; function TName.GetEmpty:Boolean; begin Result:=F=''; end; var n:TName; begin n:=TName.Create; WriteLn(n.IsEmpty); n.Value:='x'; WriteLn(n.IsEmpty); end."#),
        &["true", "false"]
    );
}

#[test]
fn property_chained_assignment_via_write() {
    assert_eq!(
        run_pascal(r#"program T; type TPair=class private FA,FB:Integer; public property A:Integer read FA write FA; property B:Integer read FB write FB; end; var p:TPair; begin p:=TPair.Create; p.A:=3; p.B:=p.A+2; WriteLn(p.B); end."#),
        &["5"]
    );
}

#[test]
fn property_notification_counter_on_write() {
    assert_eq!(
        run_pascal(r#"program T; type TWatch=class private FV,FChanges:Integer; procedure SetV(v:Integer); public property V:Integer read FV write SetV; property Changes:Integer read FChanges; end; procedure TWatch.SetV(v:Integer); begin FV:=v; Inc(FChanges); end; var w:TWatch; begin w:=TWatch.Create; w.V:=1; w.V:=2; WriteLn(w.Changes); end."#),
        &["2"]
    );
}

#[test]
fn property_indexed_negative_offset_in_getter() {
    assert_eq!(
        run_pascal(r#"program T; type TOffset=class private F:array[1..3] of Integer; function GetAt(i:Integer):Integer; public property At[i:Integer]:Integer read GetAt; end; function TOffset.GetAt(i:Integer):Integer; begin Result:=F[i+1]; end; var o:TOffset; begin o:=TOffset.Create; o.F[2]:=99; WriteLn(o.At[1]); end."#),
        &["99"]
    );
}

#[test]
fn property_default_with_count_and_add() {
    assert_eq!(
        run_pascal(r#"program T; type TStack=class private F:array[0..9] of Integer; FTop:Integer; function GetItem(i:Integer):Integer; public property Items[i:Integer]:Integer read GetItem; default; property Size:Integer read FTop; procedure Push(v:Integer); begin Inc(FTop); F[FTop]:=v; end; end; function TStack.GetItem(i:Integer):Integer; begin Result:=F[i]; end; var s:TStack; begin s:=TStack.Create; s.Push(8); WriteLn(s[0]); WriteLn(s.Size); end."#),
        &["8", "1"]
    );
}

#[test]
fn property_interface_style_on_class() {
    assert_eq!(
        run_pascal(r#"program T; type IReadable=interface function ReadVal:Integer; end; TCell=class(TInterfacedObject,IReadable) private F:Integer; public property Val:Integer read F write F; function ReadVal:Integer; begin Result:=Val; end; end; var c:IReadable; begin c:=TCell.Create; TCell(c).Val:=6; WriteLn(c.ReadVal); end."#),
        &["6"]
    );
}

#[test]
fn property_override_getter_in_descendant() {
    assert_eq!(
        run_pascal(r#"program T; type TBase=class private F:Integer; protected function GetMsg:string; virtual; public property Msg:string read GetMsg; end; TChild=class(TBase) function GetMsg:string; override; end; function TBase.GetMsg:string; begin Result:='base'; end; function TChild.GetMsg:string; begin Result:='child'; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Msg); end."#),
        &["child"]
    );
}

#[test]
fn property_array_of_char_indexed() {
    assert_eq!(
        run_pascal(r#"program T; type TChars=class private F:array[0..2] of Char; function GetCh(i:Integer):Char; procedure SetCh(i:Integer; c:Char); public property Ch[i:Integer]:Char read GetCh write SetCh; end; function TChars.GetCh(i:Integer):Char; begin Result:=F[i]; end; procedure TChars.SetCh(i:Integer; c:Char); begin F[i]:=c; end; var c:TChars; begin c:=TChars.Create; c.Ch[0]:='A'; c.Ch[1]:='B'; WriteLn(c.Ch[0]); WriteLn(c.Ch[1]); end."#),
        &["A", "B"]
    );
}

#[test]
fn property_dynamic_grow_on_setter() {
    assert_eq!(
        run_pascal(r#"program T; type TList=class private F:array of Integer; procedure SetAt(i,v:Integer); function GetAt(i:Integer):Integer; public property At[i:Integer]:Integer read GetAt write SetAt; end; procedure TList.SetAt(i,v:Integer); begin if Length(F)<=i then SetLength(F,i+1); F[i]:=v; end; function TList.GetAt(i:Integer):Integer; begin Result:=F[i]; end; var L:TList; begin L:=TList.Create; L.At[2]:=77; WriteLn(L.At[2]); end."#),
        &["77"]
    );
}

#[test]
fn property_read_in_for_loop() {
    assert_eq!(
        run_pascal(r#"program T; type TSeq=class private F:array[0..2] of Integer; function GetItem(i:Integer):Integer; public property Items[i:Integer]:Integer read GetItem; end; function TSeq.GetItem(i:Integer):Integer; begin Result:=F[i]; end; var s:TSeq; i,sum:Integer; begin s:=TSeq.Create; s.F[0]:=1; s.F[1]:=2; s.F[2]:=3; sum:=0; for i:=0 to 2 do sum:=sum+s.Items[i]; WriteLn(sum); end."#),
        &["6"]
    );
}

#[test]
fn property_class_property_integer_counter() {
    assert_eq!(
        run_pascal(r#"program T; type TGen=class strict private class var FSerial:Integer; class function Next:Integer; public class property Serial:Integer read Next; end; class function TGen.Next:Integer; begin Inc(FSerial); Result:=FSerial; end; begin WriteLn(TGen.Serial); WriteLn(TGen.Serial); end."#),
        &["1", "2"]
    );
}

#[test]
fn property_implemented_in_descendant_only() {
    assert_eq!(
        run_pascal(r#"program T; type TBase=class public function Tag:string; virtual; abstract; end; TChild=class(TBase) function Tag:string; override; end; function TChild.Tag:string; begin Result:='child'; end; var c:TChild; begin c:=TChild.Create; WriteLn(c.Tag); end."#),
        &["child"]
    );
}

#[test]
fn property_multiple_in_same_class_independent() {
    assert_eq!(
        run_pascal(r#"program T; type TPoint=class private FX,FY:Integer; public property X:Integer read FX write FX; property Y:Integer read FY write FY; end; var p:TPoint; begin p:=TPoint.Create; p.X:=2; p.Y:=3; WriteLn(p.X+p.Y); end."#),
        &["5"]
    );
}

#[test]
fn property_getter_returns_string_from_int() {
    assert_eq!(
        run_pascal(r#"program T; type TNum=class private FN:Integer; function GetHex:string; public property N:Integer read FN write FN; property Hex:string read GetHex; end; function TNum.GetHex:string; begin Result:=IntToHex(FN,2); end; var n:TNum; begin n:=TNum.Create; n.N:=255; WriteLn(n.Hex); end."#),
        &["FF"]
    );
}

#[test]
fn property_write_triggers_dependent_read() {
    assert_eq!(
        run_pascal(r#"program T; type TC=class private FR,FC:Integer; procedure SetR(v:Integer); function GetC:Integer; public property R:Integer read FR write SetR; property C:Integer read GetC; end; procedure TC.SetR(v:Integer); begin FR:=v; FC:=v*2; end; function TC.GetC:Integer; begin Result:=FC; end; var c:TC; begin c:=TC.Create; c.R:=5; WriteLn(c.C); end."#),
        &["10"]
    );
}

#[test]
fn property_indexed_string_builder() {
    assert_eq!(
        run_pascal(r#"program T; type TParts=class private F:array[0..1] of string; function GetPart(i:Integer):string; procedure SetPart(i:Integer; const s:string); public property Part[i:Integer]:string read GetPart write SetPart; function Join:string; var r:string; begin r:=Part[0]+Part[1]; Result:=r; end; end; function TParts.GetPart(i:Integer):string; begin Result:=F[i]; end; procedure TParts.SetPart(i:Integer; const s:string); begin F[i]:=s; end; var p:TParts; begin p:=TParts.Create; p.Part[0]:='ab'; p.Part[1]:='cd'; WriteLn(p.Join); end."#),
        &["abcd"]
    );
}

#[test]
fn property_default_readonly_array_element() {
    assert_eq!(
        run_pascal(r#"program T; type TView=class private F:array[0..1] of Integer; function GetV(i:Integer):Integer; public property V[i:Integer]:Integer read GetV; default; end; function TView.GetV(i:Integer):Integer; begin Result:=F[i]; end; var v:TView; begin v:=TView.Create; v.F[1]:=33; WriteLn(v[1]); end."#),
        &["33"]
    );
}

#[test]
fn property_inherited_simple_field() {
    assert_eq!(
        run_pascal(r#"program T; type TBase=class protected F:Integer; public property N:Integer read F write F; end; TChild=class(TBase); var c:TChild; begin c:=TChild.Create; c.N:=18; WriteLn(c.N); end."#),
        &["18"]
    );
}

#[test]
fn property_setter_rejects_odd_numbers() {
    assert_eq!(
        run_pascal(r#"program T; type TEven=class private F:Integer; procedure SetE(v:Integer); public property Even:Integer read F write SetE; end; procedure TEven.SetE(v:Integer); begin if v mod 2=0 then F:=v; end; var e:TEven; begin e:=TEven.Create; e.Even:=3; WriteLn(e.Even); e.Even:=4; WriteLn(e.Even); end."#),
        &["0", "4"]
    );
}

#[test]
fn property_getter_lazy_init() {
    assert_eq!(
        run_pascal(r#"program T; type TLazy=class private FReady:Boolean; FVal:Integer; function GetVal:Integer; public property Val:Integer read GetVal; end; function TLazy.GetVal:Integer; begin if not FReady then begin FVal:=99; FReady:=true; end; Result:=FVal; end; var L:TLazy; begin L:=TLazy.Create; WriteLn(L.Val); end."#),
        &["99"]
    );
}

#[test]
fn property_bool_toggle_via_setter() {
    assert_eq!(
        run_pascal(r#"program T; type TSwitch=class private F:Boolean; procedure Flip; public property On:Boolean read F; procedure Toggle; begin Flip; end; end; procedure TSwitch.Flip; begin F:=not F; end; var s:TSwitch; begin s:=TSwitch.Create; WriteLn(s.On); s.Toggle; WriteLn(s.On); end."#),
        &["false", "true"]
    );
}

#[test]
fn property_dynamic_high_index_read() {
    assert_eq!(
        run_pascal(r#"program T; type TV=class private F:array of Integer; function GetLast:Integer; public property Last:Integer read GetLast; procedure Add(v:Integer); var n:Integer; begin n:=Length(F); SetLength(F,n+1); F[n]:=v; end; end; function TV.GetLast:Integer; begin Result:=F[High(F)]; end; var v:TV; begin v:=TV.Create; v.Add(1); v.Add(9); WriteLn(v.Last); end."#),
        &["9"]
    );
}

#[test]
fn property_class_name_string_property() {
    assert_eq!(
        run_pascal(r#"program T; type TMeta=class strict private class var FTag:string; class function GetTag:string; class procedure SetTag(const s:string); public class property Tag:string read GetTag write SetTag; end; class function TMeta.GetTag:string; begin Result:=FTag; end; class procedure TMeta.SetTag(const s:string); begin FTag:=s; end; begin TMeta.Tag:='pascal'; WriteLn(TMeta.Tag); end."#),
        &["pascal"]
    );
}

#[test]
fn property_indexed_write_then_sum() {
    assert_eq!(
        run_pascal(r#"program T; type TAcc=class private F:array[0..2] of Integer; function GetA(i:Integer):Integer; procedure SetA(i:Integer; v:Integer); public property A[i:Integer]:Integer read GetA write SetA; function Sum:Integer; var i,s:Integer; begin s:=0; for i:=0 to 2 do s:=s+A[i]; Result:=s; end; end; function TAcc.GetA(i:Integer):Integer; begin Result:=F[i]; end; procedure TAcc.SetA(i:Integer; v:Integer); begin F[i]:=v; end; var a:TAcc; begin a:=TAcc.Create; a.A[0]:=1; a.A[1]:=2; a.A[2]:=3; WriteLn(a.Sum); end."#),
        &["6"]
    );
}

#[test]
fn property_read_after_free_not_used_keeps_value() {
    assert_eq!(
        run_pascal(r#"program T; type TBox=class public property Val:Integer read FVal write FVal; FVal:Integer; end; var b:TBox; begin b:=TBox.Create; b.Val:=7; WriteLn(b.Val); b.Free; end."#),
        &["7"]
    );
}

#[test]
fn property_nested_class_access() {
    assert_eq!(
        run_pascal(r#"program T; type TInner=class public property V:Integer read FV write FV; FV:Integer; end; TOuter=class public Inner:TInner; constructor Create; end; constructor TOuter.Create; begin Inner:=TInner.Create; end; var o:TOuter; begin o:=TOuter.Create; o.Inner.V:=12; WriteLn(o.Inner.V); end."#),
        &["12"]
    );
}

#[test]
fn property_getter_uses_length_of_string() {
    assert_eq!(
        run_pascal(r#"program T; type TS=class private FS:string; function GetLen:Integer; public property Text:string read FS write FS; property Len:Integer read GetLen; end; function TS.GetLen:Integer; begin Result:=Length(FS); end; var s:TS; begin s:=TS.Create; s.Text:='four'; WriteLn(s.Len); end."#),
        &["4"]
    );
}

#[test]
fn property_write_string_trim_on_set() {
    assert_eq!(
        run_pascal(r#"program T; type TTrim=class private FS:string; procedure SetS(const v:string); public property S:string read FS write SetS; end; procedure TTrim.SetS(const v:string); begin FS:=Trim(v); end; var t:TTrim; begin t:=TTrim.Create; t.S:='  ok  '; WriteLn(t.S); end."#),
        &["ok"]
    );
}

#[test]
fn property_computed_parity_boolean() {
    assert_eq!(
        run_pascal(r#"program T; type TPar=class private FN:Integer; function GetOdd:Boolean; public property N:Integer read FN write FN; property IsOdd:Boolean read GetOdd; end; function TPar.GetOdd:Boolean; begin Result:=FN mod 2<>0; end; var p:TPar; begin p:=TPar.Create; p.N:=5; WriteLn(p.IsOdd); p.N:=6; WriteLn(p.IsOdd); end."#),
        &["true", "false"]
    );
}
