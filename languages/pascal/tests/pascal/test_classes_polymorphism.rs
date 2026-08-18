/// Virtual, abstract, and dynamic dispatch scenarios beyond basic inheritance tests.
use super::helpers::run_pascal;

#[test]
fn virtual_greet_dispatches_on_dog_instance() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAnimal = class
  public function Greet: String; virtual;
end;
TDog = class(TAnimal)
  public function Greet: String; override;
end;
function TAnimal.Greet: String; begin Result := 'animal'; end;
function TDog.Greet: String; begin Result := 'woof'; end;
var d: TDog;
begin d := TDog.Create; WriteLn(d.Greet); d.Free; end."#
        ),
        &["woof"]
    );
}

#[test]
fn base_variable_calls_child_virtual_area() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TShape = class public function Area: Integer; virtual; end;
TCircle = class(TShape) public function Area: Integer; override; end;
function TShape.Area: Integer; begin Result := 0; end;
function TCircle.Area: Integer; begin Result := 314; end;
var s: TShape;
begin s := TCircle.Create; WriteLn(s.Area); s.Free; end."#
        ),
        &["314"]
    );
}

#[test]
fn abstract_draw_implemented_by_canvas() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDrawable = class abstract public function Draw: String; virtual; abstract; end;
TCanvas = class(TDrawable) public function Draw: String; override; end;
function TCanvas.Draw: String; begin Result := 'pixel'; end;
var d: TDrawable;
begin d := TCanvas.Create; WriteLn(d.Draw); d.Free; end."#
        ),
        &["pixel"]
    );
}

#[test]
fn three_level_virtual_chain_picks_leaf() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TA = class public function V: Integer; virtual; end;
TB = class(TA) public function V: Integer; override; end;
TC = class(TB) public function V: Integer; override; end;
function TA.V: Integer; begin Result := 1; end;
function TB.V: Integer; begin Result := 2; end;
function TC.V: Integer; begin Result := 3; end;
var o: TA;
begin o := TC.Create; WriteLn(o.V); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn virtual_method_uses_inherited_in_override() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Calc(n: Integer): Integer; virtual; end;
TDerived = class(TBase) public function Calc(n: Integer): Integer; override; end;
function TBase.Calc(n: Integer): Integer; begin Result := n; end;
function TDerived.Calc(n: Integer): Integer; begin Result := inherited Calc(n) * 2; end;
var d: TDerived;
begin d := TDerived.Create; WriteLn(d.Calc(5)); d.Free; end."#
        ),
        &["10"]
    );
}

#[test]
fn sibling_virtual_methods_stay_independent() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Tag: String; virtual; end;
TA = class(TBase) public function Tag: String; override; end;
TB = class(TBase) public function Tag: String; override; end;
function TBase.Tag: String; begin Result := 'base'; end;
function TA.Tag: String; begin Result := 'A'; end;
function TB.Tag: String; begin Result := 'B'; end;
var a: TA; b: TB;
begin
  a := TA.Create; b := TB.Create;
  WriteLn(a.Tag); WriteLn(b.Tag);
  a.Free; b.Free;
end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn virtual_constructor_sets_field_for_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TWidget = class
  private FId: Integer;
  public constructor Create(id: Integer); function Id: Integer; virtual;
end;
constructor TWidget.Create(id: Integer); begin FId := id; end;
function TWidget.Id: Integer; begin Result := FId; end;
var w: TWidget;
begin w := TWidget.Create(7); WriteLn(w.Id); w.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn abstract_base_with_two_concrete_implementations() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TEncoder = class abstract public function Encode(n: Integer): String; virtual; abstract; end;
THex = class(TEncoder) public function Encode(n: Integer): String; override; end;
TDec = class(TEncoder) public function Encode(n: Integer): String; override; end;
function THex.Encode(n: Integer): String; begin Result := 'h'; end;
function TDec.Encode(n: Integer): String; begin Result := IntToStr(n); end;
var a, b: TEncoder;
begin
  a := THex.Create; b := TDec.Create;
  WriteLn(a.Encode(10)); WriteLn(b.Encode(10));
  a.Free; b.Free;
end."#
        ),
        &["h", "10"]
    );
}

#[test]
fn virtual_destructor_runs_on_free() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class
  public class var Count: Integer;
  destructor Destroy; virtual;
end;
class var TBase.Count: Integer;
destructor TBase.Destroy; begin Inc(Count); inherited; end;
var o: TBase;
begin TBase.Count := 0; o := TBase.Create; o.Free; WriteLn(TBase.Count); end."#
        ),
        &["1"]
    );
}

#[test]
fn dynamic_dispatch_through_array_of_base() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TNode = class public function Val: Integer; virtual; end;
TA = class(TNode) public function Val: Integer; override; end;
TB = class(TNode) public function Val: Integer; override; end;
function TNode.Val: Integer; begin Result := 0; end;
function TA.Val: Integer; begin Result := 1; end;
function TB.Val: Integer; begin Result := 2; end;
var nodes: array[0..1] of TNode;
begin
  nodes[0] := TA.Create; nodes[1] := TB.Create;
  WriteLn(nodes[0].Val + nodes[1].Val);
  nodes[0].Free; nodes[1].Free;
end."#
        ),
        &["3"]
    );
}

#[test]
fn virtual_method_not_overridden_uses_base() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Name: String; virtual; end;
TChild = class(TBase) public function Extra: Integer; end;
function TBase.Name: String; begin Result := 'base'; end;
function TChild.Extra: Integer; begin Result := 5; end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.Name); c.Free; end."#
        ),
        &["base"]
    );
}

#[test]
fn abstract_template_method_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TJob = class abstract
  public function Run: Integer; virtual;
  function Step: Integer; virtual; abstract;
end;
TAdd = class(TJob) public function Step: Integer; override; end;
function TJob.Run: Integer; begin Result := Step + 1; end;
function TAdd.Step: Integer; begin Result := 4; end;
var j: TJob;
begin j := TAdd.Create; WriteLn(j.Run); j.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn virtual_bool_flag_in_hierarchy() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFlag = class public function On: Boolean; virtual; end;
TOn = class(TFlag) public function On: Boolean; override; end;
function TFlag.On: Boolean; begin Result := false; end;
function TOn.On: Boolean; begin Result := true; end;
var f: TFlag;
begin f := TOn.Create; WriteLn(f.On); f.Free; end."#
        ),
        &["True"]
    );
}

#[test]
fn grandchild_skips_parent_override() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TA = class public function N: Integer; virtual; end;
TB = class(TA) public function N: Integer; override; end;
TC = class(TB) end;
function TA.N: Integer; begin Result := 1; end;
function TB.N: Integer; begin Result := 2; end;
var o: TA;
begin o := TC.Create; WriteLn(o.N); o.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn virtual_string_builder_chain() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBuilder = class public function Text: String; virtual; end;
TPrefix = class(TBuilder) public function Text: String; override; end;
function TBuilder.Text: String; begin Result := 'x'; end;
function TPrefix.Text: String; begin Result := 'pre-' + inherited Text; end;
var b: TBuilder;
begin b := TPrefix.Create; WriteLn(b.Text); b.Free; end."#
        ),
        &["pre-x"]
    );
}

#[test]
fn abstract_property_via_virtual_getter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAccount = class abstract public function Balance: Integer; virtual; abstract; end;
TCash = class(TAccount) private FBal: Integer; public constructor Create(v: Integer); function Balance: Integer; override; end;
constructor TCash.Create(v: Integer); begin FBal := v; end;
function TCash.Balance: Integer; begin Result := FBal; end;
var a: TAccount;
begin a := TCash.Create(50); WriteLn(a.Balance); a.Free; end."#
        ),
        &["50"]
    );
}

#[test]
fn virtual_method_called_from_base_helper() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class
  public function Core: Integer; virtual;
  function Wrap: Integer;
end;
TChild = class(TBase) public function Core: Integer; override; end;
function TBase.Core: Integer; begin Result := 1; end;
function TChild.Core: Integer; begin Result := 9; end;
function TBase.Wrap: Integer; begin Result := Core * 2; end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.Wrap); c.Free; end."#
        ),
        &["18"]
    );
}

#[test]
fn dynamic_type_check_with_virtual_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Kind: Integer; virtual; end;
TA = class(TBase) public function Kind: Integer; override; end;
function TBase.Kind: Integer; begin Result := 0; end;
function TA.Kind: Integer; begin Result := 1; end;
var b: TBase;
begin
  b := TA.Create;
  WriteLn(b.Kind);
  b.Free;
end."#
        ),
        &["1"]
    );
}

#[test]
fn virtual_method_in_interface_style_hierarchy() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TService = class public function Ping: String; virtual; end;
TMock = class(TService) public function Ping: String; override; end;
function TService.Ping: String; begin Result := 'svc'; end;
function TMock.Ping: String; begin Result := 'mock'; end;
var s: TService;
begin s := TMock.Create; WriteLn(s.Ping); s.Free; end."#
        ),
        &["mock"]
    );
}

#[test]
fn abstract_class_two_step_construction() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class abstract
  protected FReady: Boolean;
  public constructor Create; function Ready: Boolean; virtual;
end;
TImpl = class(TBase) public constructor Create; override; end;
constructor TBase.Create; begin FReady := false; end;
function TBase.Ready: Boolean; begin Result := FReady; end;
constructor TImpl.Create; begin inherited; FReady := true; end;
var o: TBase;
begin o := TImpl.Create; WriteLn(o.Ready); o.Free; end."#
        ),
        &["True"]
    );
}

#[test]
fn virtual_dispatch_after_field_mutation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCell = class
  private FV: Integer;
  public constructor Create(v: Integer); function Read: Integer; virtual;
end;
TInc = class(TCell) public function Read: Integer; override; end;
constructor TCell.Create(v: Integer); begin FV := v; end;
function TCell.Read: Integer; begin Result := FV; end;
function TInc.Read: Integer; begin Result := inherited Read + 1; end;
var c: TCell;
begin c := TInc.Create(4); WriteLn(c.Read); c.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn override_changes_virtual_behavior_only() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class
  public function A: Integer; virtual; function B: Integer;
end;
TChild = class(TBase) public function A: Integer; override; end;
function TBase.A: Integer; begin Result := 1; end;
function TBase.B: Integer; begin Result := 2; end;
function TChild.A: Integer; begin Result := 10; end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.A); WriteLn(c.B); c.Free; end."#
        ),
        &["10", "2"]
    );
}

#[test]
fn virtual_method_returns_string_from_child() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLang = class public function Hello: String; virtual; end;
TEn = class(TLang) public function Hello: String; override; end;
TFr = class(TLang) public function Hello: String; override; end;
function TLang.Hello: String; begin Result := '?'; end;
function TEn.Hello: String; begin Result := 'hi'; end;
function TFr.Hello: String; begin Result := 'salut'; end;
var e: TLang; f: TLang;
begin
  e := TEn.Create; f := TFr.Create;
  WriteLn(e.Hello); WriteLn(f.Hello);
  e.Free; f.Free;
end."#
        ),
        &["hi", "salut"]
    );
}

#[test]
fn abstract_requires_child_for_instantiation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TParser = class abstract public function Parse(const s: String): Integer; virtual; abstract; end;
TInt = class(TParser) public function Parse(const s: String): Integer; override; end;
function TInt.Parse(const s: String): Integer; begin Result := StrToInt(s); end;
var p: TParser;
begin p := TInt.Create; WriteLn(p.Parse('12')); p.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn virtual_counter_increments_on_each_call() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class
  private FCount: Integer;
  public constructor Create; function Tick: Integer; virtual;
end;
TFast = class(TBase) public function Tick: Integer; override; end;
constructor TBase.Create; begin FCount := 0; end;
function TBase.Tick: Integer; begin Inc(FCount); Result := FCount; end;
function TFast.Tick: Integer; begin Inc(FCount, 2); Result := FCount; end;
var f: TBase;
begin f := TFast.Create; WriteLn(f.Tick); WriteLn(f.Tick); f.Free; end."#
        ),
        &["2", "4"]
    );
}

#[test]
fn base_pointer_virtual_equals_child_override() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Code: Char; virtual; end;
TChild = class(TBase) public function Code: Char; override; end;
function TBase.Code: Char; begin Result := 'B'; end;
function TChild.Code: Char; begin Result := 'C'; end;
var b: TBase;
begin b := TChild.Create; WriteLn(b.Code); b.Free; end."#
        ),
        &["C"]
    );
}

#[test]
fn virtual_method_with_multiple_overrides_in_family() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRoot = class public function Depth: Integer; virtual; end;
TMid = class(TRoot) public function Depth: Integer; override; end;
TLeaf = class(TMid) public function Depth: Integer; override; end;
function TRoot.Depth: Integer; begin Result := 1; end;
function TMid.Depth: Integer; begin Result := 2; end;
function TLeaf.Depth: Integer; begin Result := 3; end;
var o: TRoot;
begin o := TLeaf.Create; WriteLn(o.Depth); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn abstract_empty_child_still_dispatches_base_virtual() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function M: Integer; virtual; end;
TEmpty = class(TBase) end;
function TBase.M: Integer; begin Result := 42; end;
var e: TEmpty;
begin e := TEmpty.Create; WriteLn(e.M); e.Free; end."#
        ),
        &["42"]
    );
}

#[test]
fn virtual_procedure_side_effect_in_child() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class
  public class var Log: String;
  procedure Run; virtual;
end;
TChild = class(TBase) procedure Run; override; end;
class var TBase.Log: String;
procedure TBase.Run; begin Log := 'base'; end;
procedure TChild.Run; begin Log := 'child'; end;
var b: TBase;
begin
  TBase.Log := '';
  b := TChild.Create;
  b.Run;
  WriteLn(TBase.Log);
  b.Free;
end."#
        ),
        &["child"]
    );
}

#[test]
fn dynamic_dispatch_with_nil_safe_free_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TObj = class public function Tag: String; virtual; end;
TItem = class(TObj) public function Tag: String; override; end;
function TObj.Tag: String; begin Result := 'obj'; end;
function TItem.Tag: String; begin Result := 'item'; end;
var o: TObj;
begin
  o := TItem.Create;
  WriteLn(o.Tag);
  o.Free;
  WriteLn('ok');
end."#
        ),
        &["item", "ok"]
    );
}

#[test]
fn virtual_compare_polymorphic_ordering() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TOrdered = class public function Key: Integer; virtual; end;
THigh = class(TOrdered) public function Key: Integer; override; end;
function TOrdered.Key: Integer; begin Result := 1; end;
function THigh.Key: Integer; begin Result := 99; end;
var o: TOrdered;
begin o := THigh.Create; WriteLn(o.Key > 10); o.Free; end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn abstract_factory_returns_concrete_via_virtual() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TProduct = class abstract public function Sku: String; virtual; abstract; end;
TBook = class(TProduct) public function Sku: String; override; end;
function TBook.Sku: String; begin Result := 'B001'; end;
function Make: TProduct;
begin Result := TBook.Create; end;
var p: TProduct;
begin p := Make; WriteLn(p.Sku); p.Free; end."#
        ),
        &["B001"]
    );
}

#[test]
fn virtual_inherited_called_twice_in_override() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Dup(n: Integer): Integer; virtual; end;
TChild = class(TBase) public function Dup(n: Integer): Integer; override; end;
function TBase.Dup(n: Integer): Integer; begin Result := n; end;
function TChild.Dup(n: Integer): Integer; begin Result := inherited Dup(n) + inherited Dup(n); end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.Dup(3)); c.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn reintroduce_vs_override_distinct_child_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Val: Integer; virtual; end;
TChild = class(TBase) public function Val: Integer; reintroduce; end;
function TBase.Val: Integer; begin Result := 1; end;
function TChild.Val: Integer; begin Result := 2; end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.Val); c.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn virtual_method_on_intermediate_base_reference() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TA = class public function F: Integer; virtual; end;
TB = class(TA) public function F: Integer; override; end;
TC = class(TB) public function F: Integer; override; end;
function TA.F: Integer; begin Result := 1; end;
function TB.F: Integer; begin Result := 2; end;
function TC.F: Integer; begin Result := 3; end;
var b: TB;
begin b := TC.Create; WriteLn(b.F); b.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn abstract_list_item_render_hook() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRow = class abstract public function Label: String; virtual; abstract; end;
TTextRow = class(TRow) private FS: String; public constructor Create(const s: String); function Label: String; override; end;
constructor TTextRow.Create(const s: String); begin FS := s; end;
function TTextRow.Label: String; begin Result := FS; end;
var r: TRow;
begin r := TTextRow.Create('row'); WriteLn(r.Label); r.Free; end."#
        ),
        &["row"]
    );
}

#[test]
fn virtual_roundtrip_through_function_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function N: Integer; virtual; end;
TChild = class(TBase) public function N: Integer; override; end;
function TBase.N: Integer; begin Result := 0; end;
function TChild.N: Integer; begin Result := 8; end;
function Read(obj: TBase): Integer; begin Result := obj.N; end;
var c: TChild;
begin c := TChild.Create; WriteLn(Read(c)); c.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn virtual_negative_override_flips_sign() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Sign(n: Integer): Integer; virtual; end;
TNeg = class(TBase) public function Sign(n: Integer): Integer; override; end;
function TBase.Sign(n: Integer): Integer; begin Result := n; end;
function TNeg.Sign(n: Integer): Integer; begin Result := -n; end;
var o: TBase;
begin o := TNeg.Create; WriteLn(o.Sign(5)); o.Free; end."#
        ),
        &["-5"]
    );
}

#[test]
fn abstract_codec_encode_decode_pair() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCodec = class abstract
  public function Encode(n: Integer): String; virtual; abstract;
  function RoundTrip(n: Integer): Integer; virtual;
end;
TEcho = class(TCodec) public function Encode(n: Integer): String; override; end;
function TEcho.Encode(n: Integer): String; begin Result := IntToStr(n); end;
function TCodec.RoundTrip(n: Integer): Integer; begin Result := StrToInt(Encode(n)); end;
var c: TCodec;
begin c := TEcho.Create; WriteLn(c.RoundTrip(21)); c.Free; end."#
        ),
        &["21"]
    );
}

#[test]
fn virtual_dispatch_after_reassign_base_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public function Id: Integer; virtual; end;
TA = class(TBase) public function Id: Integer; override; end;
TB = class(TBase) public function Id: Integer; override; end;
function TBase.Id: Integer; begin Result := 0; end;
function TA.Id: Integer; begin Result := 1; end;
function TB.Id: Integer; begin Result := 2; end;
var b: TBase;
begin
  b := TA.Create; WriteLn(b.Id); b.Free;
  b := TB.Create; WriteLn(b.Id); b.Free;
end."#
        ),
        &["1", "2"]
    );
}
