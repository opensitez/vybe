/// Tests for advanced Pascal class features: properties, virtual/override,
/// records, with statements, class variables, destructor.

use super::helpers::run_pascal;

// ===================================================================
// PROPERTIES (GETTERS / SETTERS)
// ===================================================================

#[test] fn property_read_write() {
    assert_eq!(run_pascal(r#"program T;
type TBox = class
  private FWidth: Integer;
  public
    constructor Create(W: Integer);
    property Width: Integer read FWidth write FWidth;
  end;
constructor TBox.Create(W: Integer); begin FWidth := W; end;
var b: TBox;
begin
  b := TBox.Create(10);
  WriteLn(b.Width);
  b.Width := 25;
  WriteLn(b.Width);
end."#), &["10", "25"]);
}

#[test] fn property_getter_method() {
    assert_eq!(run_pascal(r#"program T;
type TItem = class
  private FName: String;
  public
    constructor Create(N: String);
    function GetName: String;
    property Name: String read GetName;
  end;
constructor TItem.Create(N: String); begin FName := N; end;
function TItem.GetName: String; begin Result := FName; end;
var it: TItem;
begin it := TItem.Create('widget'); WriteLn(it.Name); end."#), &["widget"]);
}

// ===================================================================
// VIRTUAL / OVERRIDE
// ===================================================================

#[test] fn virtual_override_method() {
    assert_eq!(run_pascal(r#"program T;
type TBase = class
  public function Greet: String; virtual;
end;
type TChild = class(TBase)
  public function Greet: String; override;
end;
function TBase.Greet: String; begin Result := 'Hello from Base'; end;
function TChild.Greet: String; begin Result := 'Hello from Child'; end;
var obj: TBase;
begin
  obj := TChild.Create;
  WriteLn(obj.Greet());
end."#), &["Hello from Child"]);
}

#[test] fn inherited_call() {
    assert_eq!(run_pascal(r#"program T;
type TAnimal = class
  public function Sound: String; virtual;
end;
type TDog = class(TAnimal)
  public function Sound: String; override;
end;
function TAnimal.Sound: String; begin Result := 'generic'; end;
function TDog.Sound: String; begin Result := inherited Sound + ' woof'; end;
var d: TDog;
begin d := TDog.Create; WriteLn(d.Sound()); end."#), &["generic woof"]);
}

// ===================================================================
// RECORDS
// ===================================================================

#[test] fn record_basic() {
    assert_eq!(run_pascal(r#"program T;
type TPoint = record
  X: Integer;
  Y: Integer;
end;
var p: TPoint;
begin
  p.X := 10;
  p.Y := 20;
  WriteLn(p.X + p.Y);
end."#), &["30"]);
}

#[test] fn record_as_param() {
    assert_eq!(run_pascal(r#"program T;
type TPoint = record X: Integer; Y: Integer; end;
function Magnitude(p: TPoint): Integer;
begin Result := p.X + p.Y; end;
var pt: TPoint;
begin
  pt.X := 3;
  pt.Y := 4;
  WriteLn(Magnitude(pt));
end."#), &["7"]);
}

// ===================================================================
// WITH STATEMENT
// ===================================================================

#[test] fn with_class_fields() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class
  public FX: Integer; FY: Integer;
  constructor Create(AX, AY: Integer);
end;
constructor TFoo.Create(AX, AY: Integer); begin FX := AX; FY := AY; end;
var f: TFoo;
begin
  f := TFoo.Create(10, 20);
  with f do WriteLn(FX + FY);
end."#), &["30"]);
}

// ===================================================================
// CONTINUE STATEMENT
// ===================================================================

#[test] fn continue_in_for() {
    assert_eq!(run_pascal(r#"program T;
var i: Integer;
begin
  for i := 1 to 5 do
  begin
    if i = 3 then continue;
    WriteLn(i);
  end;
end."#), &["1", "2", "4", "5"]);
}

#[test] fn continue_in_while() {
    assert_eq!(run_pascal(r#"program T;
var i: Integer;
begin
  i := 0;
  while i < 5 do
  begin
    i := i + 1;
    if i = 3 then continue;
    WriteLn(i);
  end;
end."#), &["1", "2", "4", "5"]);
}

// ===================================================================
// CLASSNAME / SIZEOF
// ===================================================================

#[test] fn classname_builtin() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class public constructor Create; end;
constructor TFoo.Create; begin end;
var f: TFoo;
begin f := TFoo.Create; WriteLn(ClassName(f)); end."#), &["tfoo"]);
}

#[test] fn sizeof_builtin() {
    assert_eq!(run_pascal("program T; begin WriteLn(SizeOf(Integer)); end."), &["4"]);
}

// ===================================================================
// MATH EXTRAS
// ===================================================================

#[test] fn math_sqrt() {
    assert_eq!(run_pascal("program T; begin WriteLn(Sqrt(16)); end."), &["4"]);
}

#[test] fn math_sqrt_real() {
    assert_eq!(run_pascal("program T; begin WriteLn(Sqrt(2.25)); end."), &["1.5"]);
}

// ===================================================================
// ARRAY OPERATIONS
// ===================================================================

#[test] fn array_append() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer;
begin
  a := [1, 2];
  Append(a, 3);
  WriteLn(Length(a));
  WriteLn(a[2]);
end."#), &["3", "3"]);
}

#[test] fn array_sort() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i: Integer;
begin
  a := [3, 1, 4, 1, 5];
  Sort(a);
  for i := 0 to High(a) do WriteLn(a[i]);
end."#), &["1", "1", "3", "4", "5"]);
}

#[test] fn array_of_strings() {
    assert_eq!(run_pascal(r#"program T;
var names: array of String; n: String;
begin
  names := ['Alice', 'Bob', 'Charlie'];
  for n in names do WriteLn(n);
end."#), &["Alice", "Bob", "Charlie"]);
}
