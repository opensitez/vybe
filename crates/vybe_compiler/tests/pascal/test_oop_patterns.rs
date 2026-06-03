/// Object-Oriented Pascal patterns: virtual/abstract, destructors, class methods,
/// method overloading, visibility, polymorphism through base-type variables.
/// Written from standard Object Pascal / Delphi conventions.
use super::helpers::run_pascal;

// ===================================================================
// DESTRUCTORS
// ===================================================================

#[test]
fn destructor_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFoo = class
  public
    constructor Create;
    destructor Destroy; override;
  end;

constructor TFoo.Create;
begin
  WriteLn('created');
end;

destructor TFoo.Destroy;
begin
  WriteLn('destroyed');
  inherited Destroy;
end;

var f: TFoo;
begin
  f := TFoo.Create;
  f.Free;
end."#
        ),
        &["created", "destroyed"]
    );
}

#[test]
fn destructor_child_calls_inherited() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    constructor Create;
    destructor Destroy; override;
  end;
  TChild = class(TBase)
  public
    constructor Create;
    destructor Destroy; override;
  end;

constructor TBase.Create; begin WriteLn('base create'); end;
destructor TBase.Destroy; begin WriteLn('base destroy'); inherited Destroy; end;
constructor TChild.Create; begin inherited Create; WriteLn('child create'); end;
destructor TChild.Destroy; begin WriteLn('child destroy'); inherited Destroy; end;

var c: TChild;
begin
  c := TChild.Create;
  c.Free;
end."#
        ),
        &[
            "base create",
            "child create",
            "child destroy",
            "base destroy"
        ]
    );
}

// ===================================================================
// VIRTUAL / OVERRIDE POLYMORPHISM
// ===================================================================

#[test]
fn virtual_method_dispatch() {
    // The fundamental OOP pattern: base-type variable, child behavior
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TShape = class
  public
    constructor Create;
    function Area: Real; virtual;
  end;
  TCircle = class(TShape)
  private
    FRadius: Real;
  public
    constructor Create(R: Real);
    function Area: Real; override;
  end;
  TSquare = class(TShape)
  private
    FSide: Real;
  public
    constructor Create(S: Real);
    function Area: Real; override;
  end;

constructor TShape.Create; begin end;
function TShape.Area: Real; begin Result := 0; end;
constructor TCircle.Create(R: Real); begin inherited Create; FRadius := R; end;
function TCircle.Area: Real; begin Result := 3.14 * FRadius * FRadius; end;
constructor TSquare.Create(S: Real); begin inherited Create; FSide := S; end;
function TSquare.Area: Real; begin Result := FSide * FSide; end;

var s: TShape;
begin
  s := TSquare.Create(5);
  WriteLn(s.Area());
end."#
        ),
        &["25"]
    );
}

#[test]
fn polymorphic_array() {
    // Array of base type holding different children
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAnimal = class
  public
    constructor Create;
    function Sound: String; virtual;
  end;
  TDog = class(TAnimal)
  public
    constructor Create;
    function Sound: String; override;
  end;
  TCat = class(TAnimal)
  public
    constructor Create;
    function Sound: String; override;
  end;

constructor TAnimal.Create; begin end;
function TAnimal.Sound: String; begin Result := '...'; end;
constructor TDog.Create; begin inherited Create; end;
function TDog.Sound: String; begin Result := 'Woof'; end;
constructor TCat.Create; begin inherited Create; end;
function TCat.Sound: String; begin Result := 'Meow'; end;

var animals: array of TAnimal;
var a: TAnimal;
begin
  animals := [TDog.Create, TCat.Create, TDog.Create];
  for a in animals do
    WriteLn(a.Sound());
end."#
        ),
        &["Woof", "Meow", "Woof"]
    );
}

// ===================================================================
// ABSTRACT METHODS
// ===================================================================

#[test]
fn abstract_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  public
    constructor Create;
    function Describe: String; virtual; abstract;
  end;
  TConcrete = class(TBase)
  public
    constructor Create;
    function Describe: String; override;
  end;

constructor TBase.Create; begin end;
constructor TConcrete.Create; begin inherited Create; end;
function TConcrete.Describe: String; begin Result := 'I am concrete'; end;

var obj: TBase;
begin
  obj := TConcrete.Create;
  WriteLn(obj.Describe());
end."#
        ),
        &["I am concrete"]
    );
}

// ===================================================================
// CLASS METHODS (STATIC)
// ===================================================================

#[test]
fn class_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMath = class
  public
    class function Add(a, b: Integer): Integer;
  end;

class function TMath.Add(a, b: Integer): Integer;
begin
  Result := a + b;
end;

begin
  WriteLn(TMath.Add(3, 4));
end."#
        ),
        &["7"]
    );
}

#[test]
fn class_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TLogger = class
  public
    class procedure Log(msg: String);
  end;

class procedure TLogger.Log(msg: String);
begin
  WriteLn('LOG: ' + msg);
end;

begin
  TLogger.Log('hello');
end."#
        ),
        &["LOG: hello"]
    );
}

// ===================================================================
// METHOD OVERLOADING
// ===================================================================

#[test]
fn method_overload_different_param_count() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCalc = class
  public
    constructor Create;
    function Add(a: Integer): Integer; overload;
    function Add(a, b: Integer): Integer; overload;
  end;

constructor TCalc.Create; begin end;
function TCalc.Add(a: Integer): Integer; begin Result := a + 10; end;
function TCalc.Add(a, b: Integer): Integer; begin Result := a + b; end;

var c: TCalc;
begin
  c := TCalc.Create;
  WriteLn(c.Add(5));
  WriteLn(c.Add(3, 4));
end."#
        ),
        &["15", "7"]
    );
}

#[test]
fn standalone_overload() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Double(x: Integer): Integer; overload;
begin Result := x * 2; end;

function Double(x: String): String; overload;
begin Result := x + x; end;

begin
  WriteLn(Double(5));
  WriteLn(Double('ab'));
end."#
        ),
        &["10", "abab"]
    );
}

// ===================================================================
// PROTECTED VISIBILITY
// ===================================================================

#[test]
fn protected_field_accessible_in_child() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBase = class
  protected
    FSecret: Integer;
  public
    constructor Create(v: Integer);
  end;
  TChild = class(TBase)
  public
    constructor Create(v: Integer);
    function GetSecret: Integer;
  end;

constructor TBase.Create(v: Integer); begin FSecret := v; end;
constructor TChild.Create(v: Integer); begin inherited Create(v); end;
function TChild.GetSecret: Integer; begin Result := FSecret; end;

var c: TChild;
begin
  c := TChild.Create(42);
  WriteLn(c.GetSecret());
end."#
        ),
        &["42"]
    );
}

// ===================================================================
// CONSTRUCTOR OVERLOADING
// ===================================================================

#[test]
fn constructor_overload() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = class
  public
    FX, FY: Integer;
    constructor Create; overload;
    constructor Create(x, y: Integer); overload;
  end;

constructor TPoint.Create;
begin FX := 0; FY := 0; end;

constructor TPoint.Create(x, y: Integer);
begin FX := x; FY := y; end;

var a, b: TPoint;
begin
  a := TPoint.Create;
  b := TPoint.Create(10, 20);
  WriteLn(a.FX); WriteLn(a.FY);
  WriteLn(b.FX); WriteLn(b.FY);
end."#
        ),
        &["0", "0", "10", "20"]
    );
}

// ===================================================================
// METHOD CALLING ANOTHER METHOD ON SELF
// ===================================================================

#[test]
fn method_calls_other_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFormatter = class
  public
    constructor Create;
    function Wrap(s: String): String;
    function FormatName(name: String): String;
  end;

constructor TFormatter.Create; begin end;
function TFormatter.Wrap(s: String): String;
begin Result := '[' + s + ']'; end;
function TFormatter.FormatName(name: String): String;
begin Result := Wrap(UpperCase(name)); end;

var f: TFormatter;
begin
  f := TFormatter.Create;
  WriteLn(f.FormatName('alice'));
end."#
        ),
        &["[ALICE]"]
    );
}

// ===================================================================
// is OPERATOR WITH INHERITANCE CHAIN
// ===================================================================

#[test]
fn is_operator_inheritance_chain() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TGrandparent = class
  public constructor Create;
  end;
  TParent = class(TGrandparent)
  public constructor Create;
  end;
  TChild = class(TParent)
  public constructor Create;
  end;

constructor TGrandparent.Create; begin end;
constructor TParent.Create; begin inherited Create; end;
constructor TChild.Create; begin inherited Create; end;

var c: TChild;
begin
  c := TChild.Create;
  if c is TChild then WriteLn('is child');
  if c is TParent then WriteLn('is parent');
  if c is TGrandparent then WriteLn('is grandparent');
end."#
        ),
        &["is child", "is parent", "is grandparent"]
    );
}

// ===================================================================
// SELF REFERENCE PASSING
// ===================================================================

#[test]
fn self_reference() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TNode = class
  public
    FValue: Integer;
    constructor Create(v: Integer);
    function GetSelf: TNode;
  end;

constructor TNode.Create(v: Integer); begin FValue := v; end;
function TNode.GetSelf: TNode; begin Result := Self; end;

var n: TNode;
begin
  n := TNode.Create(42);
  WriteLn(n.GetSelf().FValue);
end."#
        ),
        &["42"]
    );
}

// ===================================================================
// OBJECT COMPOSITION
// ===================================================================

#[test]
fn composition_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TEngine = class
  public
    FHorsepower: Integer;
    constructor Create(hp: Integer);
    function Describe: String;
  end;
  TCar = class
  public
    FName: String;
    FEngine: TEngine;
    constructor Create(name: String; hp: Integer);
    function Info: String;
  end;

constructor TEngine.Create(hp: Integer); begin FHorsepower := hp; end;
function TEngine.Describe: String; begin Result := IntToStr(FHorsepower) + 'hp'; end;
constructor TCar.Create(name: String; hp: Integer);
begin FName := name; FEngine := TEngine.Create(hp); end;
function TCar.Info: String;
begin Result := FName + ' with ' + FEngine.Describe(); end;

var car: TCar;
begin
  car := TCar.Create('Sedan', 200);
  WriteLn(car.Info());
end."#
        ),
        &["Sedan with 200hp"]
    );
}

// ===================================================================
// FACTORY PATTERN WITH CLASS FUNCTION
// ===================================================================

#[test]
fn factory_class_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TShape = class
  public
    FKind: String;
    constructor Create(k: String);
    class function Circle: TShape;
    class function Square: TShape;
  end;

constructor TShape.Create(k: String); begin FKind := k; end;
class function TShape.Circle: TShape; begin Result := TShape.Create('circle'); end;
class function TShape.Square: TShape; begin Result := TShape.Create('square'); end;

var s: TShape;
begin
  s := TShape.Circle;
  WriteLn(s.FKind);
  s := TShape.Square;
  WriteLn(s.FKind);
end."#
        ),
        &["circle", "square"]
    );
}
