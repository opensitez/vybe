use super::helpers::run_pascal;

#[test]
fn test_as_operator_downcast() {
    let src = r#"
program T;
type
  TAnimal = class
    function Speak: string; virtual;
  end;
  TDog = class(TAnimal)
    function Speak: string; override;
    function Fetch: string;
  end;

function TAnimal.Speak: string;
begin
  Result := '...';
end;

function TDog.Speak: string;
begin
  Result := 'woof';
end;

function TDog.Fetch: string;
begin
  Result := 'fetched!';
end;

var
  a: TAnimal;
  d: TDog;
begin
  d := TDog.Create;
  a := d;
  WriteLn(a.Speak);
  d := TDog(a);
  WriteLn(d.Fetch);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["woof", "fetched!"]);
}

#[test]
fn test_is_operator_base_check() {
    let src = r#"
program T;
type
  TBase = class end;
  TChild = class(TBase) end;
var
  b: TBase;
  c: TChild;
begin
  b := TBase.Create;
  c := TChild.Create;
  WriteLn(b is TBase);
  WriteLn(c is TBase);
  WriteLn(c is TChild);
  WriteLn(b is TChild);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "true", "true", "false"]);
}

#[test]
fn test_is_operator_three_level() {
    let src = r#"
program T;
type
  TL1 = class end;
  TL2 = class(TL1) end;
  TL3 = class(TL2) end;
var
  obj: TL3;
begin
  obj := TL3.Create;
  WriteLn(obj is TL1);
  WriteLn(obj is TL2);
  WriteLn(obj is TL3);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "true", "true"]);
}

#[test]
fn test_classname_method() {
    let src = r#"
program T;
type
  TMyClass = class end;
  TSubClass = class(TMyClass) end;
var
  a: TMyClass;
  b: TSubClass;
begin
  a := TMyClass.Create;
  b := TSubClass.Create;
  WriteLn(a.ClassName);
  WriteLn(b.ClassName);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["TMyClass", "TSubClass"]);
}

#[test]
fn test_inherits_from_basic() {
    let src = r#"
program T;
type
  TBase = class end;
  TChild = class(TBase) end;
var
  c: TChild;
begin
  c := TChild.Create;
  WriteLn(c.InheritsFrom(TBase));
  WriteLn(c.InheritsFrom(TChild));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn test_classtype_comparison() {
    let src = r#"
program T;
type
  TFoo = class end;
  TBar = class end;
var
  f: TFoo;
  b: TBar;
begin
  f := TFoo.Create;
  b := TBar.Create;
  WriteLn(f.ClassType = TFoo);
  WriteLn(b.ClassType = TFoo);
  WriteLn(b.ClassType = TBar);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "false", "true"]);
}

#[test]
fn test_as_then_call_method() {
    let src = r#"
program T;
type
  TShape = class
    function Area: Integer; virtual;
  end;
  TSquare = class(TShape)
    FSide: Integer;
    function Area: Integer; override;
  end;

function TShape.Area: Integer;
begin
  Result := 0;
end;

function TSquare.Area: Integer;
begin
  Result := FSide * FSide;
end;

var
  s: TShape;
  sq: TSquare;
begin
  sq := TSquare.Create;
  sq.FSide := 5;
  s := sq;
  WriteLn(s.Area);
  sq := TSquare(s);
  WriteLn(sq.Area);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["25", "25"]);
}

#[test]
fn test_polymorphic_dispatch_via_base_ref() {
    let src = r#"
program T;
type
  TProcessor = class
    function Process(n: Integer): Integer; virtual;
  end;
  TDouble = class(TProcessor)
    function Process(n: Integer): Integer; override;
  end;
  TSquare2 = class(TProcessor)
    function Process(n: Integer): Integer; override;
  end;

function TProcessor.Process(n: Integer): Integer;
begin
  Result := n;
end;

function TDouble.Process(n: Integer): Integer;
begin
  Result := n * 2;
end;

function TSquare2.Process(n: Integer): Integer;
begin
  Result := n * n;
end;

var
  processors: array[0..2] of TProcessor;
  i: Integer;
begin
  processors[0] := TProcessor.Create;
  processors[1] := TDouble.Create;
  processors[2] := TSquare2.Create;
  for i := 0 to 2 do
    WriteLn(processors[i].Process(4));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["4", "8", "16"]);
}

#[test]
fn test_class_reference_variable() {
    let src = r#"
program T;
type
  TBase = class
    class function TypeName: string; virtual;
  end;
  TDerivedA = class(TBase)
    class function TypeName: string; override;
  end;
  TDerivedB = class(TBase)
    class function TypeName: string; override;
  end;

class function TBase.TypeName: string;
begin
  Result := 'Base';
end;

class function TDerivedA.TypeName: string;
begin
  Result := 'DerivedA';
end;

class function TDerivedB.TypeName: string;
begin
  Result := 'DerivedB';
end;

begin
  WriteLn(TBase.TypeName);
  WriteLn(TDerivedA.TypeName);
  WriteLn(TDerivedB.TypeName);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Base", "DerivedA", "DerivedB"]);
}

#[test]
fn test_is_operator_with_nil() {
    let src = r#"
program T;
type
  TFoo = class end;
var
  obj: TFoo;
begin
  obj := nil;
  if obj = nil then
    WriteLn('nil object')
  else
    WriteLn('not nil');
  obj := TFoo.Create;
  if obj <> nil then
    WriteLn('created')
  else
    WriteLn('still nil');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["nil object", "created"]);
}

#[test]
fn test_class_array_polymorphism() {
    let src = r#"
program T;
type
  TLogger = class
    procedure Log(msg: string); virtual;
  end;
  TConsoleLogger = class(TLogger)
    procedure Log(msg: string); override;
  end;
  TFileLogger = class(TLogger)
    procedure Log(msg: string); override;
  end;

procedure TLogger.Log(msg: string);
begin
  WriteLn('base:' + msg);
end;

procedure TConsoleLogger.Log(msg: string);
begin
  WriteLn('console:' + msg);
end;

procedure TFileLogger.Log(msg: string);
begin
  WriteLn('file:' + msg);
end;

var
  loggers: array[0..2] of TLogger;
  i: Integer;
begin
  loggers[0] := TConsoleLogger.Create;
  loggers[1] := TFileLogger.Create;
  loggers[2] := TConsoleLogger.Create;
  for i := 0 to 2 do
    loggers[i].Log('test');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["console:test", "file:test", "console:test"]);
}

#[test]
fn test_dynamic_dispatch_no_vtable_skip() {
    let src = r#"
program T;
type
  TCalc = class
    function Compute(x: Integer): Integer; virtual;
  end;
  TAdd10 = class(TCalc)
    function Compute(x: Integer): Integer; override;
  end;
  TMul2 = class(TCalc)
    function Compute(x: Integer): Integer; override;
  end;

function TCalc.Compute(x: Integer): Integer;
begin
  Result := x;
end;

function TAdd10.Compute(x: Integer): Integer;
begin
  Result := x + 10;
end;

function TMul2.Compute(x: Integer): Integer;
begin
  Result := x * 2;
end;

function ApplyAll(calcs: array of TCalc; n, val: Integer): Integer;
var
  i: Integer;
begin
  Result := val;
  for i := 0 to n - 1 do
    Result := calcs[i].Compute(Result);
end;

var
  pipeline: array[0..1] of TCalc;
begin
  pipeline[0] := TAdd10.Create;
  pipeline[1] := TMul2.Create;
  WriteLn(ApplyAll(pipeline, 2, 5));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_rtti_classname_in_array() {
    let src = r#"
program T;
type
  TBase = class end;
  TFoo = class(TBase) end;
  TBar = class(TBase) end;
var
  objs: array[0..2] of TBase;
  i: Integer;
begin
  objs[0] := TBase.Create;
  objs[1] := TFoo.Create;
  objs[2] := TBar.Create;
  for i := 0 to 2 do
    WriteLn(objs[i].ClassName);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["TBase", "TFoo", "TBar"]);
}

#[test]
fn test_is_dispatch_conditional() {
    let src = r#"
program T;
type
  TAnimal = class
    function Name: string; virtual;
  end;
  TDog = class(TAnimal)
    function Name: string; override;
  end;
  TCat = class(TAnimal)
    function Name: string; override;
  end;

function TAnimal.Name: string;
begin
  Result := 'animal';
end;

function TDog.Name: string;
begin
  Result := 'dog';
end;

function TCat.Name: string;
begin
  Result := 'cat';
end;

procedure Greet(a: TAnimal);
begin
  if a is TDog then WriteLn('Good dog!')
  else if a is TCat then WriteLn('Nice cat!')
  else WriteLn('Hi animal!');
end;

var
  d: TDog;
  c: TCat;
  a: TAnimal;
begin
  d := TDog.Create;
  c := TCat.Create;
  a := TAnimal.Create;
  Greet(d);
  Greet(c);
  Greet(a);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Good dog!", "Nice cat!", "Hi animal!"]);
}

#[test]
fn test_inherited_method_call() {
    let src = r#"
program T;
type
  TBase = class
    function Greet: string; virtual;
  end;
  TChild = class(TBase)
    function Greet: string; override;
  end;

function TBase.Greet: string;
begin
  Result := 'Hello from base';
end;

function TChild.Greet: string;
begin
  Result := inherited Greet + ' and child';
end;

var
  c: TChild;
begin
  c := TChild.Create;
  WriteLn(c.Greet);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Hello from base and child"]);
}

#[test]
fn test_inherited_chain_three_levels() {
    let src = r#"
program T;
type
  TL1 = class
    function Tag: string; virtual;
  end;
  TL2 = class(TL1)
    function Tag: string; override;
  end;
  TL3 = class(TL2)
    function Tag: string; override;
  end;

function TL1.Tag: string;
begin
  Result := 'L1';
end;

function TL2.Tag: string;
begin
  Result := inherited Tag + '+L2';
end;

function TL3.Tag: string;
begin
  Result := inherited Tag + '+L3';
end;

var
  obj: TL3;
begin
  obj := TL3.Create;
  WriteLn(obj.Tag);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["L1+L2+L3"]);
}

#[test]
fn test_abstract_class_dispatch() {
    let src = r#"
program T;
type
  TSerializer = class
    function Serialize(v: Integer): string; virtual; abstract;
  end;
  TJsonSerializer = class(TSerializer)
    function Serialize(v: Integer): string; override;
  end;
  TCsvSerializer = class(TSerializer)
    function Serialize(v: Integer): string; override;
  end;

function TJsonSerializer.Serialize(v: Integer): string;
begin
  Result := '{"v":' + IntToStr(v) + '}';
end;

function TCsvSerializer.Serialize(v: Integer): string;
begin
  Result := IntToStr(v);
end;

procedure Output(s: TSerializer; v: Integer);
begin
  WriteLn(s.Serialize(v));
end;

var
  json: TJsonSerializer;
  csv: TCsvSerializer;
begin
  json := TJsonSerializer.Create;
  csv := TCsvSerializer.Create;
  Output(json, 42);
  Output(csv, 42);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["{\"v\":42}", "42"]);
}

#[test]
fn test_upcast_downcast_roundtrip() {
    let src = r#"
program T;
type
  TBase = class
    FID: Integer;
  end;
  TChild = class(TBase)
    FName: string;
  end;
var
  child: TChild;
  base: TBase;
  back: TChild;
begin
  child := TChild.Create;
  child.FID := 1;
  child.FName := 'test';
  base := child;
  WriteLn(base.FID);
  back := TChild(base);
  WriteLn(back.FName);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "test"]);
}

#[test]
fn test_method_override_return_type() {
    let src = r#"
program T;
type
  TConverter = class
    function Convert(s: string): string; virtual;
  end;
  TUpperConverter = class(TConverter)
    function Convert(s: string): string; override;
  end;
  TReverseConverter = class(TConverter)
    function Convert(s: string): string; override;
  end;

function TConverter.Convert(s: string): string;
begin
  Result := s;
end;

function TUpperConverter.Convert(s: string): string;
begin
  Result := UpperCase(s);
end;

function TReverseConverter.Convert(s: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := Length(s) downto 1 do
    Result := Result + s[i];
end;

var
  convs: array[0..1] of TConverter;
  i: Integer;
begin
  convs[0] := TUpperConverter.Create;
  convs[1] := TReverseConverter.Create;
  for i := 0 to 1 do
    WriteLn(convs[i].Convert('hello'));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["HELLO", "olleh"]);
}

#[test]
fn test_virtual_not_overridden_uses_base() {
    let src = r#"
program T;
type
  TBase = class
    function Name: string; virtual;
    function Full: string;
  end;
  TChild = class(TBase)
  end;

function TBase.Name: string;
begin
  Result := 'Base';
end;

function TBase.Full: string;
begin
  Result := 'Type:' + Name;
end;

var
  b: TBase;
  c: TChild;
begin
  b := TBase.Create;
  c := TChild.Create;
  WriteLn(b.Full);
  WriteLn(c.Full);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Type:Base", "Type:Base"]);
}
