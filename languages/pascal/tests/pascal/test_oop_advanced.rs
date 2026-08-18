/// Advanced OOP: constructors, destructors, class methods, visibility.
use super::helpers::run_pascal;

#[test]
fn constructor_initializes_private_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox=class
  private FV:Integer;
  public constructor Create(v:Integer); function Get:Integer;
end;
constructor TBox.Create(v:Integer); begin FV:=v; end;
function TBox.Get:Integer; begin Result:=FV; end;
var b:TBox; begin b:=TBox.Create(8); WriteLn(b.Get); b.Free; end."#
        ),
        &["8"]
    );
}

#[test]
fn destructor_runs_on_free() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TObj=class
  public class var Hit:Integer;
  destructor Destroy; override;
end;
class var TObj.Hit:Integer;
destructor TObj.Destroy; begin Inc(Hit); inherited; end;
var o:TObj; begin TObj.Hit:=0; o:=TObj.Create; o.Free; WriteLn(TObj.Hit); end."#
        ),
        &["1"]
    );
}

#[test]
fn class_method_factory_returns_instance() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPoint=class
  public FX,FY:Integer;
  class function Make(x,y:Integer):TPoint;
end;
class function TPoint.Make(x,y:Integer):TPoint;
begin Result:=TPoint.Create; Result.FX:=x; Result.FY:=y; end;
var p:TPoint; begin p:=TPoint.Make(2,3); WriteLn(p.FX+p.FY); p.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn private_field_not_accessible_use_getter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TSecret=class
  private FV:Integer;
  public constructor Create(v:Integer); function Value:Integer;
end;
constructor TSecret.Create(v:Integer); begin FV:=v; end;
function TSecret.Value:Integer; begin Result:=FV; end;
var s:TSecret; begin s:=TSecret.Create(12); WriteLn(s.Value); s.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn protected_field_visible_in_descendant() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class protected FV:Integer; public constructor Create(v:Integer); end;
TChild=class(TBase) public function Read:Integer; end;
constructor TBase.Create(v:Integer); begin FV:=v; end;
function TChild.Read:Integer; begin Result:=FV; end;
var c:TChild; begin c:=TChild.Create(6); WriteLn(c.Read); c.Free; end."#
        ),
        &["6"]
    );
}

#[test]
fn constructor_calls_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public constructor Create; end;
TChild=class(TBase) public constructor Create; end;
constructor TBase.Create; begin WriteLn('base'); end;
constructor TChild.Create; begin inherited; WriteLn('child'); end;
var c:TChild; begin c:=TChild.Create; c.Free; end."#
        ),
        &["base", "child"]
    );
}

#[test]
fn destructor_chain_child_then_base() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public destructor Destroy; virtual; end;
TChild=class(TBase) public destructor Destroy; override; end;
destructor TBase.Destroy; begin WriteLn('base'); inherited; end;
destructor TChild.Destroy; begin WriteLn('child'); inherited; end;
var c:TChild; begin c:=TChild.Create; c.Free; end."#
        ),
        &["child", "base"]
    );
}

#[test]
fn class_var_shared_across_instances() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCounted=class public class var N:Integer; constructor Create; end;
class var TCounted.N:Integer;
constructor TCounted.Create; begin Inc(N); end;
var a,b:TCounted; begin TCounted.N:=0; a:=TCounted.Create; b:=TCounted.Create; WriteLn(TCounted.N); a.Free; b.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn class_procedure_resets_counter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TReset=class public class var N:Integer; class procedure Clear; end;
class var TReset.N:Integer;
class procedure TReset.Clear; begin N:=0; end;
begin TReset.N:=5; TReset.Clear; WriteLn(TReset.N); end."#
        ),
        &["0"]
    );
}

#[test]
fn public_method_calls_private_helper() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCalc=class
  private function Double(n:Integer):Integer;
  public function Quad(n:Integer):Integer;
end;
function TCalc.Double(n:Integer):Integer; begin Result:=n*2; end;
function TCalc.Quad(n:Integer):Integer; begin Result:=Double(Double(n)); end;
var c:TCalc; begin c:=TCalc.Create; WriteLn(c.Quad(3)); c.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn constructor_with_multiple_params() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRect=class public W,H:Integer; constructor Create(w,h:Integer); end;
constructor TRect.Create(w,h:Integer); begin W:=w; H:=h; end;
var r:TRect; begin r:=TRect.Create(4,5); WriteLn(r.W*r.H); r.Free; end."#
        ),
        &["20"]
    );
}

#[test]
fn destructor_frees_owned_string_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TMsg=class public FText:String; constructor Create(const s:String); end;
constructor TMsg.Create(const s:String); begin FText:=s; end;
var m:TMsg; begin m:=TMsg.Create('bye'); WriteLn(m.FText); m.Free; WriteLn('ok'); end."#
        ),
        &["bye", "ok"]
    );
}

#[test]
fn class_function_returns_constant() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TVersion=class public class function Major:Integer; end;
class function TVersion.Major:Integer; begin Result:=1; end;
begin WriteLn(TVersion.Major); end."#
        ),
        &["1"]
    );
}

#[test]
fn strict_private_method_via_public_wrapper() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TGate=class
  strict private function Core:Integer;
  public function Open:Integer;
end;
function TGate.Core:Integer; begin Result:=7; end;
function TGate.Open:Integer; begin Result:=Core; end;
var g:TGate; begin g:=TGate.Create; WriteLn(g.Open); g.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn published_property_style_getter_setter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TProp=class
  private FV:Integer;
  public function GetV:Integer; procedure SetV(v:Integer); property Value:Integer read GetV write SetV;
end;
function TProp.GetV:Integer; begin Result:=FV; end;
procedure TProp.SetV(v:Integer); begin FV:=v; end;
var p:TProp; begin p:=TProp.Create; p.Value:=9; WriteLn(p.Value); p.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn constructor_raises_field_before_use() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInit=class public FReady:Boolean; constructor Create; end;
constructor TInit.Create; begin FReady:=true; end;
var o:TInit; begin o:=TInit.Create; WriteLn(o.FReady); o.Free; end."#
        ),
        &["True"]
    );
}

#[test]
fn class_destructor_pattern_on_free() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRes=class public class var Freed:Boolean; destructor Destroy; override; end;
class var TRes.Freed:Boolean;
destructor TRes.Destroy; begin Freed:=true; inherited; end;
var r:TRes; begin TRes.Freed:=false; r:=TRes.Create; r.Free; WriteLn(TRes.Freed); end."#
        ),
        &["True"]
    );
}

#[test]
fn nested_class_method_calls_instance_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair=class
  public FA,FB:Integer;
  class function Sum(a,b:Integer):Integer;
  function Total:Integer;
end;
class function TPair.Sum(a,b:Integer):Integer; begin Result:=a+b; end;
function TPair.Total:Integer; begin Result:=Sum(FA,FB); end;
var p:TPair; begin p:=TPair.Create; p.FA:=2; p.FB:=3; WriteLn(p.Total); p.Free; end."#
        ),
        &["5"]
    );
}

#[test]
fn visibility_public_method_on_private_object_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type THidden=class
  private FX:Integer;
  public procedure SetX(v:Integer); function GetX:Integer;
end;
procedure THidden.SetX(v:Integer); begin FX:=v; end;
function THidden.GetX:Integer; begin Result:=FX; end;
var h:THidden; begin h:=THidden.Create; h.SetX(4); WriteLn(h.GetX); h.Free; end."#
        ),
        &["4"]
    );
}

#[test]
fn constructor_allocates_dynamic_array_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBag=class public Items:array of Integer; constructor Create; end;
constructor TBag.Create; begin SetLength(Items,2); Items[0]:=1; Items[1]:=2; end;
var b:TBag; begin b:=TBag.Create; WriteLn(b.Items[1]); b.Free; end."#
        ),
        &["2"]
    );
}

#[test]
fn destructor_child_logs_before_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public class var Log:String; destructor Destroy; virtual; end;
TChild=class(TBase) destructor Destroy; override; end;
class var TBase.Log:String;
destructor TBase.Destroy; begin Log:=Log+'B'; inherited; end;
destructor TChild.Destroy; begin Log:=Log+'C'; inherited; end;
var c:TChild; begin TBase.Log:=''; c:=TChild.Create; c.Free; WriteLn(TBase.Log); end."#
        ),
        &["CB"]
    );
}

#[test]
fn class_method_creates_named_singleton() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TSing=class public class var Inst:TSing; class function Instance:TSing; end;
class var TSing.Inst:TSing;
class function TSing.Instance:TSing; begin if Inst=nil then Inst:=TSing.Create; Result:=Inst; end;
var a,b:TSing; begin a:=TSing.Instance; b:=TSing.Instance; WriteLn(a=b); a.Free; TSing.Inst:=nil; end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn private_constructor_enforced_via_class_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TOnly=class
  strict private constructor Create(v:Integer);
  public class function Make(v:Integer):TOnly;
  public FV:Integer;
end;
constructor TOnly.Create(v:Integer); begin FV:=v; end;
class function TOnly.Make(v:Integer):TOnly; begin Result:=TOnly.Create(v); end;
var o:TOnly; begin o:=TOnly.Make(3); WriteLn(o.FV); o.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn protected_constructor_in_family() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class protected constructor Create(v:Integer); public FV:Integer; end;
TChild=class(TBase) public constructor Create(v:Integer); end;
constructor TBase.Create(v:Integer); begin FV:=v; end;
constructor TChild.Create(v:Integer); begin inherited(v); end;
var c:TChild; begin c:=TChild.Create(11); WriteLn(c.FV); c.Free; end."#
        ),
        &["11"]
    );
}

#[test]
fn class_property_style_accessor() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TMeta=class public class var FName:String; class function Name:String; end;
class var TMeta.FName:String;
class function TMeta.Name:String; begin Result:=FName; end;
begin TMeta.FName:='vybe'; WriteLn(TMeta.Name); end."#
        ),
        &["vybe"]
    );
}

#[test]
fn constructor_virtual_method_ready() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public constructor Create; function Tag:String; virtual; end;
TChild=class(TBase) public function Tag:String; override; end;
constructor TBase.Create; begin end;
function TBase.Tag:String; begin Result:='base'; end;
function TChild.Tag:String; begin Result:='child'; end;
var c:TChild; begin c:=TChild.Create; WriteLn(c.Tag); c.Free; end."#
        ),
        &["child"]
    );
}

#[test]
fn destructor_virtual_in_hierarchy() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public class var N:Integer; destructor Destroy; virtual; end;
TChild=class(TBase) destructor Destroy; override; end;
class var TBase.N:Integer;
destructor TBase.Destroy; begin Inc(N); inherited; end;
destructor TChild.Destroy; begin inherited; end;
var c:TChild; begin TBase.N:=0; c:=TChild.Create; c.Free; WriteLn(TBase.N); end."#
        ),
        &["1"]
    );
}

#[test]
fn public_inherits_private_data_via_descendant() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class private FV:Integer; protected procedure SetV(v:Integer); public function GetV:Integer; end;
TChild=class(TBase) public procedure Bump; end;
procedure TBase.SetV(v:Integer); begin FV:=v; end;
function TBase.GetV:Integer; begin Result:=FV; end;
procedure TChild.Bump; begin SetV(GetV+1); end;
var c:TChild; begin c:=TChild.Create; c.Bump; WriteLn(c.GetV); c.Free; end."#
        ),
        &["1"]
    );
}

#[test]
fn class_method_string_builder() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TStr=class public class function Join(a,b:String):String; end;
class function TStr.Join(a,b:String):String; begin Result:=a+'+'+b; end;
begin WriteLn(TStr.Join('1','2')); end."#
        ),
        &["1+2"]
    );
}

#[test]
fn constructor_exception_safe_field_init() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TSafe=class public FOk:Boolean; constructor Create; end;
constructor TSafe.Create; begin FOk:=true; end;
var s:TSafe; begin s:=TSafe.Create; WriteLn(s.FOk); s.Free; end."#
        ),
        &["True"]
    );
}

#[test]
fn multiple_instances_distinct_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TItem=class public V:Integer; constructor Create(v:Integer); end;
constructor TItem.Create(v:Integer); begin V:=v; end;
var a,b:TItem; begin a:=TItem.Create(1); b:=TItem.Create(2); WriteLn(a.V+b.V); a.Free; b.Free; end."#
        ),
        &["3"]
    );
}

#[test]
fn class_procedure_increments_global() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TGlobal=class public class var Hits:Integer; class procedure Touch; end;
class var TGlobal.Hits:Integer;
class procedure TGlobal.Touch; begin Inc(Hits); end;
begin TGlobal.Hits:=0; TGlobal.Touch; TGlobal.Touch; WriteLn(TGlobal.Hits); end."#
        ),
        &["2"]
    );
}

#[test]
fn strict_protected_visible_in_child_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class strict protected FV:Integer; public constructor Create(v:Integer); end;
TChild=class(TBase) public function Read:Integer; end;
constructor TBase.Create(v:Integer); begin FV:=v; end;
function TChild.Read:Integer; begin Result:=FV; end;
var c:TChild; begin c:=TChild.Create(15); WriteLn(c.Read); c.Free; end."#
        ),
        &["15"]
    );
}

#[test]
fn constructor_sets_string_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TName=class public FName:String; constructor Create(const n:String); end;
constructor TName.Create(const n:String); begin FName:=n; end;
var x:TName; begin x:=TName.Create('ann'); WriteLn(x.FName); x.Free; end."#
        ),
        &["ann"]
    );
}

#[test]
fn destructor_clears_class_var_flag() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFlag=class public class var Alive:Boolean; constructor Create; destructor Destroy; override; end;
class var TFlag.Alive:Boolean;
constructor TFlag.Create; begin Alive:=true; end;
destructor TFlag.Destroy; begin Alive:=false; inherited; end;
var f:TFlag; begin f:=TFlag.Create; f.Free; WriteLn(TFlag.Alive); end."#
        ),
        &["False"]
    );
}

#[test]
fn class_function_math_helper() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TMath=class public class function Square(n:Integer):Integer; end;
class function TMath.Square(n:Integer):Integer; begin Result:=n*n; end;
begin WriteLn(TMath.Square(6)); end."#
        ),
        &["36"]
    );
}

#[test]
fn private_section_multiple_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair=class
  private FA,FB:Integer;
  public constructor Create(a,b:Integer); function Sum:Integer;
end;
constructor TPair.Create(a,b:Integer); begin FA:=a; FB:=b; end;
function TPair.Sum:Integer; begin Result:=FA+FB; end;
var p:TPair; begin p:=TPair.Create(3,4); WriteLn(p.Sum); p.Free; end."#
        ),
        &["7"]
    );
}

#[test]
fn constructor_overload_by_param_count() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TVal=class public V:Integer; constructor Create; overload; constructor Create(v:Integer); overload; end;
constructor TVal.Create; begin V:=0; end;
constructor TVal.Create(v:Integer); begin V:=v; end;
var a,b:TVal; begin a:=TVal.Create; b:=TVal.Create(5); WriteLn(a.V); WriteLn(b.V); a.Free; b.Free; end."#
        ),
        &["0", "5"]
    );
}

#[test]
fn public_class_method_hides_instance_complexity() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFactory=class
  private constructor Create(v:Integer);
  public class function Build(v:Integer):TFactory;
  public FV:Integer;
end;
constructor TFactory.Create(v:Integer); begin FV:=v; end;
class function TFactory.Build(v:Integer):TFactory; begin Result:=TFactory.Create(v); end;
var f:TFactory; begin f:=TFactory.Build(99); WriteLn(f.FV); f.Free; end."#
        ),
        &["99"]
    );
}

#[test]
fn destructor_inherited_called_explicitly() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase=class public class var Seq:String; destructor Destroy; virtual; end;
TChild=class(TBase) destructor Destroy; override; end;
class var TBase.Seq:String;
destructor TBase.Destroy; begin Seq:=Seq+'B'; inherited; end;
destructor TChild.Destroy; begin Seq:=Seq+'C'; inherited; end;
var c:TChild; begin TBase.Seq:=''; c:=TChild.Create; c.Free; WriteLn(TBase.Seq); end."#
        ),
        &["CB"]
    );
}
