/// Tests for advanced Pascal class features: multi-level inheritance,
/// class variables, destructors, method patterns, polymorphism.

use super::helpers::run_pascal;

// ===================================================================
// MULTI-LEVEL INHERITANCE
// ===================================================================

#[test] fn inheritance_three_levels() {
    assert_eq!(run_pascal(r#"program T;
type
  TAnimal = class
    public function Kind: String; virtual;
  end;
  TMammal = class(TAnimal)
    public function Kind: String; override;
  end;
  TDog = class(TMammal)
    public function Kind: String; override;
  end;
function TAnimal.Kind: String; begin Result := 'Animal'; end;
function TMammal.Kind: String; begin Result := 'Mammal'; end;
function TDog.Kind: String; begin Result := 'Dog'; end;
var d: TDog;
begin
  d := TDog.Create;
  WriteLn(d.Kind());
end."#), &["Dog"]);
}

#[test] fn inherited_field_access() {
    assert_eq!(run_pascal(r#"program T;
type
  TBase = class
    public FX: Integer;
    constructor Create(X: Integer);
  end;
  TChild = class(TBase)
    public FY: Integer;
    constructor Create(X, Y: Integer);
  end;
constructor TBase.Create(X: Integer); begin FX := X; end;
constructor TChild.Create(X, Y: Integer);
begin
  inherited Create(X);
  FY := Y;
end;
var c: TChild;
begin
  c := TChild.Create(10, 20);
  WriteLn(c.FX);
  WriteLn(c.FY);
end."#), &["10", "20"]);
}

#[test] fn inherited_method_from_grandparent() {
    assert_eq!(run_pascal(r#"program T;
type
  TBase = class
    public function Hello: String;
  end;
  TMid = class(TBase)
  end;
  TLeaf = class(TMid)
  end;
function TBase.Hello: String; begin Result := 'hi'; end;
var obj: TLeaf;
begin
  obj := TLeaf.Create;
  WriteLn(obj.Hello());
end."#), &["hi"]);
}

// ===================================================================
// CONSTRUCTOR WITH INHERITED CALL
// ===================================================================

#[test] fn constructor_inherited_create() {
    assert_eq!(run_pascal(r#"program T;
type
  TShape = class
    public FColor: String;
    constructor Create(C: String);
  end;
  TCircle = class(TShape)
    public FRadius: Integer;
    constructor Create(C: String; R: Integer);
  end;
constructor TShape.Create(C: String); begin FColor := C; end;
constructor TCircle.Create(C: String; R: Integer);
begin
  inherited Create(C);
  FRadius := R;
end;
var c: TCircle;
begin
  c := TCircle.Create('red', 5);
  WriteLn(c.FColor);
  WriteLn(c.FRadius);
end."#), &["red", "5"]);
}

// ===================================================================
// CLASS METHODS CALLING EACH OTHER
// ===================================================================

#[test] fn method_calls_method() {
    assert_eq!(run_pascal(r#"program T;
type TCalc = class
  public
    FVal: Integer;
    constructor Create(V: Integer);
    function Double: Integer;
    function Quadruple: Integer;
end;
constructor TCalc.Create(V: Integer); begin FVal := V; end;
function TCalc.Double: Integer; begin Result := FVal * 2; end;
function TCalc.Quadruple: Integer; begin Result := Double() * 2; end;
var c: TCalc;
begin
  c := TCalc.Create(5);
  WriteLn(c.Quadruple());
end."#), &["20"]);
}

// ===================================================================
// CONSTRUCTOR WITHOUT PARAMS (ZERO-ARG)
// ===================================================================

#[test] fn zero_arg_constructor_no_parens() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class
  public FVal: Integer;
  constructor Create;
end;
constructor TFoo.Create; begin FVal := 42; end;
var f: TFoo;
begin
  f := TFoo.Create;
  WriteLn(f.FVal);
end."#), &["42"]);
}

// ===================================================================
// MULTIPLE CLASSES IN ONE PROGRAM
// ===================================================================

#[test] fn multiple_classes() {
    assert_eq!(run_pascal(r#"program T;
type
  TPoint = class
    public FX, FY: Integer;
    constructor Create(X, Y: Integer);
  end;
  TRect = class
    public FW, FH: Integer;
    constructor Create(W, H: Integer);
    function Area: Integer;
  end;
constructor TPoint.Create(X, Y: Integer); begin FX := X; FY := Y; end;
constructor TRect.Create(W, H: Integer); begin FW := W; FH := H; end;
function TRect.Area: Integer; begin Result := FW * FH; end;
var p: TPoint; r: TRect;
begin
  p := TPoint.Create(3, 4);
  r := TRect.Create(10, 5);
  WriteLn(p.FX + p.FY);
  WriteLn(r.Area());
end."#), &["7", "50"]);
}

// ===================================================================
// CLASS WITH PROCEDURE (SUB) AND FUNCTION
// ===================================================================

#[test] fn class_procedure_and_function() {
    assert_eq!(run_pascal(r#"program T;
type TCounter = class
  public
    FCount: Integer;
    constructor Create;
    procedure Increment;
    function GetCount: Integer;
end;
constructor TCounter.Create; begin FCount := 0; end;
procedure TCounter.Increment; begin FCount := FCount + 1; end;
function TCounter.GetCount: Integer; begin Result := FCount; end;
var c: TCounter;
begin
  c := TCounter.Create;
  c.Increment;
  c.Increment;
  c.Increment;
  WriteLn(c.GetCount());
end."#), &["3"]);
}

// ===================================================================
// IS / AS WITH CLASSES
// ===================================================================

#[test] fn is_operator_inheritance() {
    assert_eq!(run_pascal(r#"program T;
type
  TBase = class end;
  TChild = class(TBase) end;
var b: TBase;
begin
  b := TChild.Create;
  if b is TChild then WriteLn('yes') else WriteLn('no');
end."#), &["yes"]);
}

// ===================================================================
// CLASS FIELDS MULTIPLE ON ONE LINE
// ===================================================================

#[test] fn fields_same_type() {
    assert_eq!(run_pascal(r#"program T;
type TPoint = class
  public FX, FY, FZ: Integer;
  constructor Create(X, Y, Z: Integer);
end;
constructor TPoint.Create(X, Y, Z: Integer);
begin FX := X; FY := Y; FZ := Z; end;
var p: TPoint;
begin
  p := TPoint.Create(1, 2, 3);
  WriteLn(p.FX + p.FY + p.FZ);
end."#), &["6"]);
}

// ===================================================================
// CLASS WITH STRING OPERATIONS
// ===================================================================

#[test] fn class_string_methods() {
    assert_eq!(run_pascal(r#"program T;
type TPerson = class
  public
    FFirst, FLast: String;
    constructor Create(F, L: String);
    function FullName: String;
    function Initials: String;
end;
constructor TPerson.Create(F, L: String); begin FFirst := F; FLast := L; end;
function TPerson.FullName: String; begin Result := FFirst + ' ' + FLast; end;
function TPerson.Initials: String; begin Result := Copy(FFirst, 1, 1) + Copy(FLast, 1, 1); end;
var p: TPerson;
begin
  p := TPerson.Create('John', 'Doe');
  WriteLn(p.FullName());
  WriteLn(p.Initials());
end."#), &["John Doe", "JD"]);
}

// ===================================================================
// CLASS ARRAY ITERATION
// ===================================================================

#[test] fn iterate_class_array() {
    assert_eq!(run_pascal(r#"program T;
type TVal = class
  public FN: Integer;
  constructor Create(N: Integer);
end;
constructor TVal.Create(N: Integer); begin FN := N; end;
var items: array of TVal; i: Integer;
begin
  items := [TVal.Create(10), TVal.Create(20), TVal.Create(30)];
  for i := 0 to High(items) do WriteLn(items[i].FN);
end."#), &["10", "20", "30"]);
}
