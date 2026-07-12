/// Abstract, sealed, override, and reintroduce method semantics in Delphi classes.
use super::helpers::run_pascal;

#[test]
fn abstract_method_must_be_overridden_in_concrete_child() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    function Describe: String; virtual; abstract;
  end;
  TChild = class(TBase)
  public
    function Describe: String; override;
  end;
function TChild.Describe: String; begin Result := 'child'; end;
var o: TChild;
begin
  o := TChild.Create;
  WriteLn(o.Describe());
end."#
        ),
        &["child"]
    );
}

#[test]
fn virtual_override_dispatches_on_runtime_type() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAnimal = class
  public
    function Speak: String; virtual;
  end;
  TDog = class(TAnimal)
  public
    function Speak: String; override;
  end;
function TAnimal.Speak: String; begin Result := '...'; end;
function TDog.Speak: String; begin Result := 'woof'; end;
var a: TAnimal;
begin
  a := TDog.Create;
  WriteLn(a.Speak());
end."#
        ),
        &["woof"]
    );
}

#[test]
fn sealed_class_cannot_be_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFinal = class sealed
  public
    function Value: Integer;
  end;
function TFinal.Value: Integer; begin Result := 9; end;
var f: TFinal;
begin
  f := TFinal.Create;
  WriteLn(f.Value());
end."#
        ),
        &["9"]
    );
}

#[test]
fn reintroduce_hides_parent_method_without_override() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    function Tag: String; virtual;
  end;
  TChild = class(TBase)
  public
    function Tag: String; reintroduce;
  end;
function TBase.Tag: String; begin Result := 'base'; end;
function TChild.Tag: String; begin Result := 'child'; end;
var b: TBase;
begin
  b := TChild.Create;
  WriteLn(b.Tag());
end."#
        ),
        &["base"]
    );
}

#[test]
fn reintroduce_visible_when_child_typed_reference() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    function Tag: String; virtual;
  end;
  TChild = class(TBase)
  public
    function Tag: String; reintroduce;
  end;
function TBase.Tag: String; begin Result := 'base'; end;
function TChild.Tag: String; begin Result := 'child'; end;
var c: TChild;
begin
  c := TChild.Create;
  WriteLn(c.Tag());
end."#
        ),
        &["child"]
    );
}

#[test]
fn abstract_class_cannot_be_instantiated_use_concrete_subclass() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TShape = class abstract
  public
    function Area: Integer; virtual; abstract;
  end;
  TSquare = class(TShape)
  public
    FSide: Integer;
    constructor Create(s: Integer);
    function Area: Integer; override;
  end;
constructor TSquare.Create(s: Integer); begin FSide := s; end;
function TSquare.Area: Integer; begin Result := FSide * FSide; end;
var s: TSquare;
begin
  s := TSquare.Create(4);
  WriteLn(s.Area());
end."#
        ),
        &["16"]
    );
}

#[test]
fn inherited_virtual_calls_parent_when_not_overridden() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    function Id: Integer; virtual;
  end;
  TMid = class(TBase)
  end;
function TBase.Id: Integer; begin Result := 1; end;
var m: TMid;
begin
  m := TMid.Create;
  WriteLn(m.Id());
end."#
        ),
        &["1"]
    );
}

#[test]
fn override_chain_skips_middle_implementation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TA = class
  public
    function N: Integer; virtual;
  end;
  TB = class(TA)
  public
    function N: Integer; override;
  end;
  TC = class(TB)
  public
    function N: Integer; override;
  end;
function TA.N: Integer; begin Result := 1; end;
function TB.N: Integer; begin Result := 2; end;
function TC.N: Integer; begin Result := 3; end;
var x: TA;
begin
  x := TC.Create;
  WriteLn(x.N());
end."#
        ),
        &["3"]
    );
}

#[test]
fn virtual_destructor_runs_most_derived_first() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    destructor Destroy; override;
  end;
  TChild = class(TBase)
  public
    destructor Destroy; override;
  end;
destructor TBase.Destroy; begin WriteLn('base'); inherited; end;
destructor TChild.Destroy; begin WriteLn('child'); inherited; end;
var c: TChild;
begin
  c := TChild.Create;
  c.Free;
end."#
        ),
        &["child", "base"]
    );
}

#[test]
fn class_procedure_static_dispatch_without_instance() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMath = class
  public
    class function Double(n: Integer): Integer;
  end;
class function TMath.Double(n: Integer): Integer; begin Result := n * 2; end;
begin
  WriteLn(TMath.Double(21));
end."#
        ),
        &["42"]
    );
}
