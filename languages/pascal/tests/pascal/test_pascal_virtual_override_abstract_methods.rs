use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 14: Virtual, Override, Abstract Methods & Polymorphism
// ═══════════════════════════════════════════════════════════

#[test]
fn test_virtual_method_overriding() {
    let out = run_pascal(
        r#"
program Test;
type TAnimal = class
  public procedure MakeSound; virtual;
end;
type TDog = class(TAnimal)
  public procedure MakeSound; override;
end;
procedure TAnimal.MakeSound; begin WriteLn('Generic'); end;
procedure TDog.MakeSound; begin WriteLn('Bark'); end;
var a: TAnimal;
begin
  a := TDog.Create;
  a.MakeSound;
  a.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Bark"]);
}

#[test]
fn test_abstract_method_invocation() {
    let out = run_pascal(
        r#"
program Test;
type TShape = class
  public function GetArea: Real; virtual; abstract;
end;
type TSquare = class(TShape)
  private FSide: Real;
  public constructor Create(S: Real);
  public function GetArea: Real; override;
end;
constructor TSquare.Create(S: Real); begin FSide := S; end;
function TSquare.GetArea: Real; begin Result := FSide * FSide; end;
var s: TShape;
begin
  s := TSquare.Create(4.0);
  WriteLn(s.GetArea);
  s.Free;
end.
"#,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn test_polymorphic_array_dispatch() {
    let out = run_pascal(
        r#"
program Test;
type TBaseWidget = class
  public procedure Render; virtual;
end;
type TButton = class(TBaseWidget)
  public procedure Render; override;
end;
type TLabel = class(TBaseWidget)
  public procedure Render; override;
end;
procedure TBaseWidget.Render; begin WriteLn('Widget'); end;
procedure TButton.Render; begin WriteLn('Button'); end;
procedure TLabel.Render; begin WriteLn('Label'); end;
var widgets: array[1..2] of TBaseWidget;
    i: Integer;
begin
  widgets[1] := TButton.Create;
  widgets[2] := TLabel.Create;
  for i := 1 to 2 do
    widgets[i].Render;
  for i := 1 to 2 do
    widgets[i].Free;
end.
"#,
    );
    assert_eq!(out, vec!["Button", "Label"]);
}

#[test]
fn test_inherited_method_call_in_override() {
    let out = run_pascal(
        r#"
program Test;
type TParent = class
  public procedure Greet; virtual;
end;
type TChild = class(TParent)
  public procedure Greet; override;
end;
procedure TParent.Greet; begin WriteLn('Hello Parent'); end;
procedure TChild.Greet;
begin
  inherited Greet;
  WriteLn('Hello Child');
end;
var c: TChild;
begin
  c := TChild.Create;
  c.Greet;
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello Parent", "Hello Child"]);
}

#[test]
fn test_three_level_override_hierarchy() {
    let out = run_pascal(
        r#"
program Test;
type TLevel1 = class
  public procedure Action; virtual;
end;
type TLevel2 = class(TLevel1)
  public procedure Action; override;
end;
type TLevel3 = class(TLevel2)
  public procedure Action; override;
end;
procedure TLevel1.Action; begin WriteLn('L1'); end;
procedure TLevel2.Action; begin WriteLn('L2'); end;
procedure TLevel3.Action; begin WriteLn('L3'); end;
var obj: TLevel1;
begin
  obj := TLevel3.Create;
  obj.Action;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["L3"]);
}

#[test]
fn test_virtual_destructor_polymorphism() {
    let out = run_pascal(
        r#"
program Test;
type TBaseResource = class
  public destructor Destroy; override;
end;
type TCustomResource = class(TBaseResource)
  public destructor Destroy; override;
end;
destructor TBaseResource.Destroy; begin WriteLn('BaseCleaned'); inherited Destroy; end;
destructor TCustomResource.Destroy; begin WriteLn('CustomCleaned'); inherited Destroy; end;
var res: TBaseResource;
begin
  res := TCustomResource.Create;
  res.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CustomCleaned", "BaseCleaned"]);
}

#[test]
fn test_dynamic_method_specifier() {
    let out = run_pascal(
        r#"
program Test;
type TDynBase = class
  public procedure Exec; dynamic;
end;
type TDynSub = class(TDynBase)
  public procedure Exec; override;
end;
procedure TDynBase.Exec; begin WriteLn('DynBase'); end;
procedure TDynSub.Exec; begin WriteLn('DynSub'); end;
var d: TDynBase;
begin
  d := TDynSub.Create;
  d.Exec;
  d.Free;
end.
"#,
    );
    assert_eq!(out, vec!["DynSub"]);
}

#[test]
fn test_virtual_method_called_by_base_method() {
    let out = run_pascal(
        r#"
program Test;
type TFramework = class
  protected procedure ProcessStep; virtual;
  public procedure RunFramework;
end;
type TAppFramework = class(TFramework)
  protected procedure ProcessStep; override;
end;
procedure TFramework.ProcessStep; begin WriteLn('DefaultStep'); end;
procedure TFramework.RunFramework;
begin
  WriteLn('Start');
  ProcessStep;
  WriteLn('End');
end;
procedure TAppFramework.ProcessStep; begin WriteLn('AppCustomStep'); end;
var app: TFramework;
begin
  app := TAppFramework.Create;
  app.RunFramework;
  app.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Start", "AppCustomStep", "End"]);
}

#[test]
fn test_abstract_class_mixed_concrete_methods() {
    let out = run_pascal(
        r#"
program Test;
type TAbstractLogger = class
  public procedure LogHeader;
  public procedure LogBody(msg: String); virtual; abstract;
end;
type TConsoleLogger = class(TAbstractLogger)
  public procedure LogBody(msg: String); override;
end;
procedure TAbstractLogger.LogHeader; begin WriteLn('=== LOG ==='); end;
procedure TConsoleLogger.LogBody(msg: String); begin WriteLn(msg); end;
var l: TAbstractLogger;
begin
  l := TConsoleLogger.Create;
  l.LogHeader;
  l.LogBody('Entry 1');
  l.Free;
end.
"#,
    );
    assert_eq!(out, vec!["=== LOG ===", "Entry 1"]);
}

#[test]
fn test_override_function_returning_record() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TProvider = class
  public function GetPosition: TPoint; virtual;
end;
type TCustomProvider = class(TProvider)
  public function GetPosition: TPoint; override;
end;
function TProvider.GetPosition: TPoint; begin Result.X := 0; Result.Y := 0; end;
function TCustomProvider.GetPosition: TPoint; begin Result.X := 100; Result.Y := 200; end;
var p: TProvider; pt: TPoint;
begin
  p := TCustomProvider.Create;
  pt := p.GetPosition;
  WriteLn(pt.X);
  WriteLn(pt.Y);
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_override_procedure_mutates_state() {
    let out = run_pascal(
        r#"
program Test;
type TCounter = class
  public Value: Integer;
  public procedure Add; virtual;
end;
type TDoubleCounter = class(TCounter)
  public procedure Add; override;
end;
procedure TCounter.Add; begin Value := Value + 1; end;
procedure TDoubleCounter.Add; begin Value := Value + 2; end;
var c: TCounter;
begin
  c := TDoubleCounter.Create;
  c.Value := 10;
  c.Add;
  WriteLn(c.Value);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn test_virtual_method_is_type_check_before_call() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class procedure Action; virtual; end;
type TSub = class(TBase) procedure Action; override; end;
procedure TBase.Action; begin WriteLn('Base'); end;
procedure TSub.Action; begin WriteLn('Sub'); end;
var b: TBase;
begin
  b := TSub.Create;
  if b is TSub then
    (b as TSub).Action;
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Sub"]);
}

#[test]
fn test_inherited_call_in_middle_override() {
    let out = run_pascal(
        r#"
program Test;
type TA = class procedure Step; virtual; end;
type TB = class(TA) procedure Step; override; end;
type TC = class(TB) procedure Step; override; end;
procedure TA.Step; begin WriteLn('A'); end;
procedure TB.Step; begin WriteLn('B'); inherited Step; end;
procedure TC.Step; begin WriteLn('C'); inherited Step; end;
var obj: TA;
begin
  obj := TC.Create;
  obj.Step;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["C", "B", "A"]);
}

#[test]
fn test_virtual_function_with_string_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TFormatter = class
  public function FormatText(s: String): String; virtual;
end;
type TUpperFormatter = class(TFormatter)
  public function FormatText(s: String): String; override;
end;
function TFormatter.FormatText(s: String): String; begin Result := s; end;
function TUpperFormatter.FormatText(s: String): String; begin Result := UpperCase(s); end;
var f: TFormatter;
begin
  f := TUpperFormatter.Create;
  WriteLn(f.FormatText('hello pascal'));
  f.Free;
end.
"#,
    );
    assert_eq!(out, vec!["HELLO PASCAL"]);
}

#[test]
fn test_abstract_method_two_subclasses() {
    let out = run_pascal(
        r#"
program Test;
type TCalc = class function Exec(a, b: Integer): Integer; virtual; abstract; end;
type TAddCalc = class(TCalc) function Exec(a, b: Integer): Integer; override; end;
type TMulCalc = class(TCalc) function Exec(a, b: Integer): Integer; override; end;
function TAddCalc.Exec(a, b: Integer): Integer; begin Result := a + b; end;
function TMulCalc.Exec(a, b: Integer): Integer; begin Result := a * b; end;
var c1, c2: TCalc;
begin
  c1 := TAddCalc.Create; c2 := TMulCalc.Create;
  WriteLn(c1.Exec(10, 5));
  WriteLn(c2.Exec(10, 5));
  c1.Free; c2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["15", "50"]);
}

#[test]
fn test_override_with_var_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TModifier = class
  public procedure Transform(var x: Integer); virtual;
end;
type TScaleModifier = class(TModifier)
  public procedure Transform(var x: Integer); override;
end;
procedure TModifier.Transform(var x: Integer); begin Inc(x); end;
procedure TScaleModifier.Transform(var x: Integer); begin x := x * 10; end;
var m: TModifier; val: Integer;
begin
  m := TScaleModifier.Create;
  val := 5;
  m.Transform(val);
  WriteLn(val);
  m.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_override_with_const_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TPrinter = class
  public procedure PrintVal(const s: String); virtual;
end;
type TPrefixPrinter = class(TPrinter)
  public procedure PrintVal(const s: String); override;
end;
procedure TPrinter.PrintVal(const s: String); begin WriteLn(s); end;
procedure TPrefixPrinter.PrintVal(const s: String); begin WriteLn('> ' + s); end;
var p: TPrinter;
begin
  p := TPrefixPrinter.Create;
  p.PrintVal('Message');
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["> Message"]);
}

#[test]
fn test_abstract_method_returning_boolean() {
    let out = run_pascal(
        r#"
program Test;
type TValidator = class
  public function IsValid(data: String): Boolean; virtual; abstract;
end;
type TNonEmptyValidator = class(TValidator)
  public function IsValid(data: String): Boolean; override;
end;
function TNonEmptyValidator.IsValid(data: String): Boolean;
begin
  Result := Length(data) > 0;
end;
var v: TValidator;
begin
  v := TNonEmptyValidator.Create;
  WriteLn(v.IsValid('test'));
  WriteLn(v.IsValid(''));
  v.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_override_method_accessing_derived_fields() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  public function GetDescription: String; virtual;
end;
type TItem = class(TBase)
  public Name: String; Price: Real;
  constructor Create(N: String; P: Real);
  public function GetDescription: String; override;
end;
constructor TItem.Create(N: String; P: Real); begin Name := N; Price := P; end;
function TBase.GetDescription: String; begin Result := 'BaseItem'; end;
function TItem.GetDescription: String; begin Result := Name + '=$' + Price.ToString; end;
var b: TBase;
begin
  b := TItem.Create('Book', 15.99);
  WriteLn(b.GetDescription);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Book=$15.99"]);
}

#[test]
fn test_abstract_class_factory_instantiation() {
    let out = run_pascal(
        r#"
program Test;
type TEngine = class
  public procedure Start; virtual; abstract;
end;
type TElectricEngine = class(TEngine)
  public procedure Start; override;
end;
function CreateEngine: TEngine;
begin
  Result := TElectricEngine.Create;
end;
procedure TElectricEngine.Start; begin WriteLn('SilentStart'); end;
var eng: TEngine;
begin
  eng := CreateEngine;
  eng.Start;
  eng.Free;
end.
"#,
    );
    assert_eq!(out, vec!["SilentStart"]);
}
