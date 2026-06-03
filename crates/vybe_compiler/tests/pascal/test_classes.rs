use super::helpers::run_pascal;

// -- Base class --

#[test]
fn class_create_and_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFoo = class public FVal: Integer; constructor Create(V: Integer); end;
constructor TFoo.Create(V: Integer); begin FVal := V; end;
var f: TFoo;
begin f := TFoo.Create(42); WriteLn(f.FVal); end."#
        ),
        &["42"]
    );
}

#[test]
fn class_method_returns_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAnimal = class
  private FName: String;
  public constructor Create(AName: String); function Speak: String;
end;
constructor TAnimal.Create(AName: String); begin FName := AName; end;
function TAnimal.Speak: String; begin Result := FName + ' speaks'; end;
var a: TAnimal;
begin a := TAnimal.Create('Rex'); WriteLn(a.Speak()); end."#
        ),
        &["Rex speaks"]
    );
}

#[test]
fn class_zero_param_constructor() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFoo = class public FX: Integer; constructor Create; end;
constructor TFoo.Create; begin FX := 99; end;
var f: TFoo; begin f := TFoo.Create; WriteLn(f.FX); end."#
        ),
        &["99"]
    );
}

#[test]
fn class_multiple_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPoint = class public FX: Integer; FY: Integer; constructor Create(AX, AY: Integer); end;
constructor TPoint.Create(AX, AY: Integer); begin FX := AX; FY := AY; end;
var p: TPoint;
begin p := TPoint.Create(10, 20); WriteLn(p.FX + p.FY); end."#
        ),
        &["30"]
    );
}

#[test]
fn class_method_modifies_state() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCounter = class public FCount: Integer;
  constructor Create; function Increment: Integer; end;
constructor TCounter.Create; begin FCount := 0; end;
function TCounter.Increment: Integer; begin FCount := FCount + 1; Result := FCount; end;
var c: TCounter;
begin c := TCounter.Create; c.Increment(); c.Increment(); c.Increment(); WriteLn(c.FCount); end."#
        ),
        &["3"]
    );
}

#[test]
fn class_multiple_methods() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCalc = class public FVal: Integer;
  constructor Create(V: Integer); function GetVal: Integer; function Add(X: Integer): Integer; end;
constructor TCalc.Create(V: Integer); begin FVal := V; end;
function TCalc.GetVal: Integer; begin Result := FVal; end;
function TCalc.Add(X: Integer): Integer; begin FVal := FVal + X; Result := FVal; end;
var c: TCalc;
begin c := TCalc.Create(10); c.Add(5); c.Add(3); WriteLn(c.GetVal()); end."#
        ),
        &["18"]
    );
}

#[test]
fn class_method_with_params() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TMath = class public
  constructor Create; function Add(a, b: Integer): Integer; end;
constructor TMath.Create; begin end;
function TMath.Add(a, b: Integer): Integer; begin Result := a + b; end;
var m: TMath;
begin m := TMath.Create; WriteLn(m.Add(3, 4)); end."#
        ),
        &["7"]
    );
}

#[test]
fn class_string_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPerson = class public FName: String; FAge: Integer;
  constructor Create(AName: String; AAge: Integer); function Desc: String; end;
constructor TPerson.Create(AName: String; AAge: Integer); begin FName := AName; FAge := AAge; end;
function TPerson.Desc: String; begin Result := FName + ' is ' + IntToStr(FAge); end;
var p: TPerson;
begin p := TPerson.Create('Alice', 30); WriteLn(p.Desc()); end."#
        ),
        &["Alice is 30"]
    );
}

// -- Multiple instances --

#[test]
fn class_two_instances_independent() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox = class public FVal: Integer;
  constructor Create(V: Integer); function GetVal: Integer; end;
constructor TBox.Create(V: Integer); begin FVal := V; end;
function TBox.GetVal: Integer; begin Result := FVal; end;
var a, b: TBox;
begin a := TBox.Create(10); b := TBox.Create(20); WriteLn(a.GetVal()); WriteLn(b.GetVal()); end."#
        ),
        &["10", "20"]
    );
}

#[test]
fn class_instances_mutate_independently() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCounter = class public FCount: Integer;
  constructor Create(V: Integer); function MyInc: Integer; function GetCount: Integer; end;
constructor TCounter.Create(V: Integer); begin FCount := V; end;
function TCounter.MyInc: Integer; begin FCount := FCount + 1; Result := FCount; end;
function TCounter.GetCount: Integer; begin Result := FCount; end;
var a, b: TCounter;
begin a := TCounter.Create(0); b := TCounter.Create(0);
  a.MyInc(); a.MyInc(); a.MyInc(); b.MyInc();
  WriteLn(a.GetCount()); WriteLn(b.GetCount()); end."#
        ),
        &["3", "1"]
    );
}

#[test]
fn class_three_instances() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TVal = class public FV: Integer; constructor Create(V: Integer); end;
constructor TVal.Create(V: Integer); begin FV := V; end;
var a, b, c: TVal;
begin a := TVal.Create(1); b := TVal.Create(2); c := TVal.Create(3);
  WriteLn(a.FV + b.FV + c.FV); end."#
        ),
        &["6"]
    );
}

// -- Inheritance --

#[test]
fn class_child_inherits_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAnimal = class private FName: String; public
  constructor Create(AName: String); function Speak: String; end;
type TDog = class(TAnimal) public constructor Create(AName: String); end;
constructor TAnimal.Create(AName: String); begin FName := AName; end;
function TAnimal.Speak: String; begin Result := FName + ' speaks'; end;
constructor TDog.Create(AName: String); begin inherited Create(AName); end;
var d: TDog;
begin d := TDog.Create('Rex'); WriteLn(d.Speak()); end."#
        ),
        &["Rex speaks"]
    );
}

#[test]
fn class_child_inherits_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public FVal: Integer; constructor Create(V: Integer); end;
type TChild = class(TBase) public constructor Create(V: Integer); end;
constructor TBase.Create(V: Integer); begin FVal := V; end;
constructor TChild.Create(V: Integer); begin inherited Create(V); end;
var c: TChild;
begin c := TChild.Create(42); WriteLn(c.FVal); end."#
        ),
        &["42"]
    );
}

#[test]
fn class_child_overrides_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public constructor Create; function Greet: String; end;
type TChild = class(TBase) public constructor Create; function Greet: String; end;
constructor TBase.Create; begin end;
function TBase.Greet: String; begin Result := 'base'; end;
constructor TChild.Create; begin inherited Create; end;
function TChild.Greet: String; begin Result := 'child'; end;
var c: TChild;
begin c := TChild.Create; WriteLn(c.Greet()); end."#
        ),
        &["child"]
    );
}

#[test]
fn class_child_adds_own_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public FX: Integer; constructor Create(X: Integer); end;
type TChild = class(TBase) public FY: Integer;
  constructor Create(X, Y: Integer); function Sum: Integer; end;
constructor TBase.Create(X: Integer); begin FX := X; end;
constructor TChild.Create(X, Y: Integer); begin inherited Create(X); FY := Y; end;
function TChild.Sum: Integer; begin Result := FX + FY; end;
var c: TChild;
begin c := TChild.Create(10, 20); WriteLn(c.Sum()); end."#
        ),
        &["30"]
    );
}

#[test]
fn class_child_adds_own_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class public FName: String; constructor Create(N: String); function GetName: String; end;
type TChild = class(TBase) public constructor Create(N: String); function Upper: String; end;
constructor TBase.Create(N: String); begin FName := N; end;
function TBase.GetName: String; begin Result := FName; end;
constructor TChild.Create(N: String); begin inherited Create(N); end;
function TChild.Upper: String; begin Result := UpperCase(FName); end;
var c: TChild;
begin c := TChild.Create('hello'); WriteLn(c.GetName()); WriteLn(c.Upper()); end."#
        ),
        &["hello", "HELLO"]
    );
}
