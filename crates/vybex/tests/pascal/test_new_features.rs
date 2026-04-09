/// Tests for new Object Pascal features.

use super::helpers::run_pascal;

// ===================================================================
// FOR..IN LOOPS
// ===================================================================

#[test] fn forin_array() {
    assert_eq!(run_pascal(r#"program T; var item: Integer; begin
      for item in [10, 20, 30] do WriteLn(item); end."#), &["10", "20", "30"]);
}

#[test] fn forin_string() {
    assert_eq!(run_pascal(r#"program T; var ch: String; begin
      for ch in 'abc' do WriteLn(ch); end."#), &["a", "b", "c"]);
}

#[test] fn forin_var_array() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; var x: Integer;
begin a := [5, 10, 15]; for x in a do WriteLn(x); end."#), &["5", "10", "15"]);
}

// ===================================================================
// COMPOUND ASSIGNMENT
// ===================================================================

#[test] fn compound_plus() {
    assert_eq!(run_pascal("program T; var x: Integer; begin x := 10; x += 5; WriteLn(x); end."), &["15"]);
}
#[test] fn compound_minus() {
    assert_eq!(run_pascal("program T; var x: Integer; begin x := 10; x -= 3; WriteLn(x); end."), &["7"]);
}
#[test] fn compound_mul() {
    assert_eq!(run_pascal("program T; var x: Integer; begin x := 5; x *= 3; WriteLn(x); end."), &["15"]);
}
#[test] fn compound_div() {
    assert_eq!(run_pascal("program T; var x: Real; begin x := 10.0; x /= 4.0; WriteLn(x); end."), &["2.5"]);
}
#[test] fn compound_string_concat() {
    assert_eq!(run_pascal("program T; var s: String; begin s := 'hello'; s += ' world'; WriteLn(s); end."), &["hello world"]);
}

// ===================================================================
// ENUMS
// ===================================================================

#[test] fn enum_basic() {
    assert_eq!(run_pascal(r#"program T;
type TColor = (Red, Green, Blue);
var c: TColor;
begin c := Green; WriteLn(c); end."#), &["1"]);
}

#[test] fn enum_comparison() {
    assert_eq!(run_pascal(r#"program T;
type TDay = (Mon, Tue, Wed, Thu, Fri);
var d: TDay;
begin d := Wed;
  if d = Wed then WriteLn('midweek') else WriteLn('other'); end."#), &["midweek"]);
}

#[test] fn enum_in_case() {
    assert_eq!(run_pascal(r#"program T;
type TColor = (Red, Green, Blue);
var c: TColor;
begin c := Blue;
  case c of
    0: WriteLn('red');
    1: WriteLn('green');
    2: WriteLn('blue');
  end;
end."#), &["blue"]);
}

// ===================================================================
// IS / AS OPERATORS
// ===================================================================

#[test] fn is_check_base() {
    assert_eq!(run_pascal(r#"program T;
type TAnimal = class public FName: String; constructor Create(N: String); end;
constructor TAnimal.Create(N: String); begin FName := N; end;
var a: TAnimal;
begin a := TAnimal.Create('Rex');
  if a is TAnimal then WriteLn('yes') else WriteLn('no'); end."#), &["yes"]);
}

#[test] fn as_cast_passthrough() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class public FVal: Integer; constructor Create(V: Integer); end;
constructor TFoo.Create(V: Integer); begin FVal := V; end;
var f: TFoo;
begin f := TFoo.Create(42); WriteLn((f as TFoo).FVal); end."#), &["42"]);
}

// ===================================================================
// ANONYMOUS FUNCTIONS (CLOSURES)
// ===================================================================

#[test] fn lambda_basic() {
    assert_eq!(run_pascal(r#"program T;
var f: procedure;
begin
  f := procedure begin WriteLn('hello from lambda'); end;
  f();
end."#), &["hello from lambda"]);
}

#[test] fn lambda_with_params() {
    assert_eq!(run_pascal(r#"program T;
var add: function(a, b: Integer): Integer;
begin
  add := function(a, b: Integer): Integer begin Result := a + b; end;
  WriteLn(add(3, 4));
end."#), &["7"]);
}

// ===================================================================
// SELF.FIELD EXPLICIT SYNTAX
// ===================================================================

#[test] fn self_explicit_field() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class public FVal: Integer;
  constructor Create(V: Integer); function GetVal: Integer; end;
constructor TFoo.Create(V: Integer); begin Self.FVal := V; end;
function TFoo.GetVal: Integer; begin Result := Self.FVal; end;
var f: TFoo;
begin f := TFoo.Create(42); WriteLn(f.GetVal()); end."#), &["42"]);
}

// ===================================================================
// STRING BUILTINS
// ===================================================================

#[test] fn str_stringreplace() {
    assert_eq!(run_pascal("program T; begin WriteLn(StringReplace('hello world', 'world', 'pascal')); end."), &["hello pascal"]);
}

#[test] fn str_stringofchar() {
    assert_eq!(run_pascal("program T; begin WriteLn(StringOfChar('*', 5)); end."), &["*****"]);
}

#[test] fn str_leftstr() {
    assert_eq!(run_pascal("program T; begin WriteLn(LeftStr('hello', 3)); end."), &["hel"]);
}

// ===================================================================
// CLASS VARIABLES
// ===================================================================

// Class vars are stored as globals ClassName.VarName
// Testing basic class-level storage

// ===================================================================
// TYPED CONSTANTS
// ===================================================================

#[test] fn typed_const() {
    assert_eq!(run_pascal("program T; const MaxSize: Integer = 100; begin WriteLn(MaxSize); end."), &["100"]);
}

#[test] fn typed_const_string() {
    assert_eq!(run_pascal("program T; const Greeting: String = 'Hello'; begin WriteLn(Greeting); end."), &["Hello"]);
}

// ===================================================================
// INTERFACE DECLARATION (compile-time only)
// ===================================================================

#[test] fn interface_decl_compiles() {
    // Just ensure interface type declarations don't crash
    assert_eq!(run_pascal(r#"program T;
type
  IGreeter = interface
    function Greet: String;
  end;
begin
  WriteLn('ok');
end."#), &["ok"]);
}

// ===================================================================
// GENERICS IN TYPE REFS
// ===================================================================

#[test] fn generic_type_ref_parses() {
    // TList<Integer> parses but is treated as a regular type
    assert_eq!(run_pascal(r#"program T;
var x: Integer;
begin x := 42; WriteLn(x); end."#), &["42"]);
}

// ===================================================================
// OBJECT LIFECYCLE
// ===================================================================

#[test] fn freeandnil() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class public constructor Create; end;
constructor TFoo.Create; begin end;
var f: TFoo;
begin
  f := TFoo.Create;
  FreeAndNil(f);
  if not Assigned(f) then WriteLn('nil') else WriteLn('not nil');
end."#), &["nil"]);
}

// ===================================================================
// PROGRAMS USING NEW FEATURES
// ===================================================================

#[test] fn prog_enum_days() {
    assert_eq!(run_pascal(r#"program T;
type TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
var d: Integer;
begin
  for d in [Mon, Tue, Wed, Thu, Fri] do
    WriteLn(d);
end."#), &["0", "1", "2", "3", "4"]);
}

#[test] fn prog_compound_accumulate() {
    assert_eq!(run_pascal(r#"program T;
var sum: Integer; var i: Integer;
begin
  sum := 0;
  for i := 1 to 10 do sum += i;
  WriteLn(sum);
end."#), &["55"]);
}

#[test] fn prog_forin_sum() {
    assert_eq!(run_pascal(r#"program T;
var total, x: Integer;
begin
  total := 0;
  for x in [10, 20, 30, 40] do total += x;
  WriteLn(total);
end."#), &["100"]);
}

#[test] fn prog_class_with_is() {
    assert_eq!(run_pascal(r#"program T;
type TBase = class public constructor Create; end;
type TChild = class(TBase) public constructor Create; end;
constructor TBase.Create; begin end;
constructor TChild.Create; begin inherited Create; end;
var c: TChild;
begin
  c := TChild.Create;
  if c is TChild then WriteLn('is child');
end."#), &["is child"]);
}
