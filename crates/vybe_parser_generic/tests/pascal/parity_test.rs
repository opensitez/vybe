/// Parity test: run EVERY source from the compiler tests through the generic parser.
/// Any failure here means a gap vs the hand-written parser.

use vybe_parser_generic::grammar::*;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;

fn grammar() -> GrammarDef {
    super::parse_tests::pascal_grammar_pub()
}

fn must_parse(src: &str) {
    let g = grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
    if let Err(e) = parse(&tokens, &g) {
        panic!("PARSE FAILED: {}\nSource:\n{}", e, src);
    }
}

// ── Every source from test_builtins.rs ──────────────────────────────────────
#[test] fn b_writeln_one()    { must_parse("program T; begin WriteLn('hello'); end."); }
#[test] fn b_writeln_multi()  { must_parse("program T; begin WriteLn('a', 'b', 'c'); end."); }
#[test] fn b_writeln_int()    { must_parse("program T; begin WriteLn(42); end."); }
#[test] fn b_writeln_real()   { must_parse("program T; begin WriteLn(3.14); end."); }
#[test] fn b_writeln_bool()   { must_parse("program T; begin WriteLn(true); end."); }
#[test] fn b_writeln_expr()   { must_parse("program T; begin WriteLn(2 + 3); end."); }
#[test] fn b_multi_writeln()  { must_parse("program T; begin WriteLn('a'); WriteLn('b'); WriteLn('c'); end."); }
#[test] fn b_length()         { must_parse("program T; begin WriteLn(Length('hello')); end."); }
#[test] fn b_length_empty()   { must_parse("program T; begin WriteLn(Length('')); end."); }
#[test] fn b_uppercase()      { must_parse("program T; begin WriteLn(UpperCase('hello')); end."); }
#[test] fn b_lowercase()      { must_parse("program T; begin WriteLn(LowerCase('HELLO')); end."); }
#[test] fn b_trim()           { must_parse("program T; begin WriteLn(Trim('  hi  ')); end."); }
#[test] fn b_concat_plus()    { must_parse("program T; begin WriteLn('foo' + 'bar'); end."); }
#[test] fn b_concat_fn()      { must_parse("program T; begin WriteLn(Concat('a', 'b', 'c')); end."); }
#[test] fn b_concat_var()     { must_parse("program T; var a, b: String; begin a := 'hello'; b := ' world'; WriteLn(a + b); end."); }
#[test] fn b_multi_concat()   { must_parse("program T; begin WriteLn('a' + 'b' + 'c' + 'd'); end."); }
#[test] fn b_inttostr()       { must_parse("program T; begin WriteLn(IntToStr(42)); end."); }
#[test] fn b_floattostr()     { must_parse("program T; begin WriteLn(FloatToStr(3.14)); end."); }
#[test] fn b_strtoint()       { must_parse("program T; begin WriteLn(StrToInt('42')); end."); }
#[test] fn b_strtofloat()     { must_parse("program T; begin WriteLn(StrToFloat('3.14')); end."); }
#[test] fn b_abs_pos()        { must_parse("program T; begin WriteLn(Abs(5)); end."); }
#[test] fn b_abs_neg()        { must_parse("program T; begin WriteLn(Abs(-5)); end."); }
#[test] fn b_abs_zero()       { must_parse("program T; begin WriteLn(Abs(0)); end."); }
#[test] fn b_sqr()            { must_parse("program T; begin WriteLn(Sqr(4)); end."); }
#[test] fn b_sqr_neg()        { must_parse("program T; begin WriteLn(Sqr(-3)); end."); }
#[test] fn b_min()            { must_parse("program T; begin WriteLn(Min(3, 7)); end."); }
#[test] fn b_max()            { must_parse("program T; begin WriteLn(Max(3, 7)); end."); }
#[test] fn b_min_eq()         { must_parse("program T; begin WriteLn(Min(5, 5)); end."); }
#[test] fn b_floor()          { must_parse("program T; begin WriteLn(Floor(3.7)); end."); }
#[test] fn b_ceil()           { must_parse("program T; begin WriteLn(Ceil(3.2)); end."); }
#[test] fn b_round()          { must_parse("program T; begin WriteLn(Round(3.5)); end."); }
#[test] fn b_trunc()          { must_parse("program T; begin WriteLn(Trunc(3.9)); end."); }
#[test] fn b_power()          { must_parse("program T; begin WriteLn(Power(2, 10)); end."); }
#[test] fn b_succ()           { must_parse("program T; begin WriteLn(Succ(5)); end."); }
#[test] fn b_pred()           { must_parse("program T; begin WriteLn(Pred(5)); end."); }
#[test] fn b_inc()            { must_parse("program T; var x: Integer; begin x := 5; Inc(x); WriteLn(x); end."); }
#[test] fn b_dec()            { must_parse("program T; var x: Integer; begin x := 5; Dec(x); WriteLn(x); end."); }
#[test] fn b_inc_multiple()   { must_parse("program T; var x: Integer; begin x := 0; Inc(x); Inc(x); Inc(x); WriteLn(x); end."); }
#[test] fn b_dec_neg()        { must_parse("program T; var x: Integer; begin x := 1; Dec(x); Dec(x); WriteLn(x); end."); }
#[test] fn b_inc_dec()        { must_parse("program T; var x: Integer; begin x := 10; Inc(x); Dec(x); WriteLn(x); end."); }
#[test] fn b_array_literal()  { must_parse("program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(a[0]); WriteLn(a[2]); end."); }
#[test] fn b_array_assign()   { must_parse("program T; var a: array of Integer; begin a := [1, 2, 3]; a[1] := 99; WriteLn(a[1]); end."); }
#[test] fn b_array_length()   { must_parse("program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(Length(a)); end."); }
#[test] fn b_array_high()     { must_parse("program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(High(a)); end."); }
#[test] fn b_array_low()      { must_parse("program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(Low(a)); end."); }
#[test] fn b_array_iterate()  { must_parse("program T; var a: array of Integer; var i: Integer; begin a := [5, 10, 15]; for i := 0 to High(a) do WriteLn(a[i]); end."); }
#[test] fn b_assigned_nil()   { must_parse("program T; begin if not Assigned(nil) then WriteLn('y'); end."); }

// ── Every source from test_classes.rs ───────────────────────────────────────
#[test] fn c_create_field()    { must_parse("program T; type TFoo = class public FVal: Integer; constructor Create(V: Integer); end; constructor TFoo.Create(V: Integer); begin FVal := V; end; var f: TFoo; begin f := TFoo.Create(42); WriteLn(f.FVal); end."); }
#[test] fn c_method_returns()  { must_parse(r#"program T; type TAnimal = class private FName: String; public constructor Create(AName: String); function Speak: String; end; constructor TAnimal.Create(AName: String); begin FName := AName; end; function TAnimal.Speak: String; begin Result := FName + ' speaks'; end; var a: TAnimal; begin a := TAnimal.Create('Rex'); WriteLn(a.Speak()); end."#); }
#[test] fn c_zero_param()      { must_parse("program T; type TFoo = class public FX: Integer; constructor Create; end; constructor TFoo.Create; begin FX := 99; end; var f: TFoo; begin f := TFoo.Create; WriteLn(f.FX); end."); }
#[test] fn c_multi_fields()    { must_parse("program T; type TPoint = class public FX: Integer; FY: Integer; constructor Create(AX, AY: Integer); end; constructor TPoint.Create(AX, AY: Integer); begin FX := AX; FY := AY; end; var p: TPoint; begin p := TPoint.Create(10, 20); WriteLn(p.FX + p.FY); end."); }
#[test] fn c_modify_state()    { must_parse(r#"program T; type TCounter = class public FCount: Integer; constructor Create; function Increment: Integer; end; constructor TCounter.Create; begin FCount := 0; end; function TCounter.Increment: Integer; begin FCount := FCount + 1; Result := FCount; end; var c: TCounter; begin c := TCounter.Create; c.Increment(); c.Increment(); c.Increment(); WriteLn(c.FCount); end."#); }
#[test] fn c_multi_methods()   { must_parse(r#"program T; type TCalc = class public FVal: Integer; constructor Create(V: Integer); function GetVal: Integer; function Add(X: Integer): Integer; end; constructor TCalc.Create(V: Integer); begin FVal := V; end; function TCalc.GetVal: Integer; begin Result := FVal; end; function TCalc.Add(X: Integer): Integer; begin FVal := FVal + X; Result := FVal; end; var c: TCalc; begin c := TCalc.Create(10); c.Add(5); c.Add(3); WriteLn(c.GetVal()); end."#); }
#[test] fn c_method_params()   { must_parse(r#"program T; type TMath = class public constructor Create; function Add(a, b: Integer): Integer; end; constructor TMath.Create; begin end; function TMath.Add(a, b: Integer): Integer; begin Result := a + b; end; var m: TMath; begin m := TMath.Create; WriteLn(m.Add(3, 4)); end."#); }
#[test] fn c_string_field()    { must_parse(r#"program T; type TPerson = class public FName: String; FAge: Integer; constructor Create(AName: String; AAge: Integer); function Desc: String; end; constructor TPerson.Create(AName: String; AAge: Integer); begin FName := AName; FAge := AAge; end; function TPerson.Desc: String; begin Result := FName + ' is ' + IntToStr(FAge); end; var p: TPerson; begin p := TPerson.Create('Alice', 30); WriteLn(p.Desc()); end."#); }
#[test] fn c_two_instances()   { must_parse(r#"program T; type TBox = class public FVal: Integer; constructor Create(V: Integer); function GetVal: Integer; end; constructor TBox.Create(V: Integer); begin FVal := V; end; function TBox.GetVal: Integer; begin Result := FVal; end; var a, b: TBox; begin a := TBox.Create(10); b := TBox.Create(20); WriteLn(a.GetVal()); WriteLn(b.GetVal()); end."#); }
#[test] fn c_inherit_method()  { must_parse(r#"program T; type TAnimal = class private FName: String; public constructor Create(AName: String); function Speak: String; end; type TDog = class(TAnimal) public constructor Create(AName: String); end; constructor TAnimal.Create(AName: String); begin FName := AName; end; function TAnimal.Speak: String; begin Result := FName + ' speaks'; end; constructor TDog.Create(AName: String); begin inherited Create(AName); end; var d: TDog; begin d := TDog.Create('Rex'); WriteLn(d.Speak()); end."#); }
#[test] fn c_inherit_field()   { must_parse(r#"program T; type TBase = class public FVal: Integer; constructor Create(V: Integer); end; type TChild = class(TBase) public constructor Create(V: Integer); end; constructor TBase.Create(V: Integer); begin FVal := V; end; constructor TChild.Create(V: Integer); begin inherited Create(V); end; var c: TChild; begin c := TChild.Create(42); WriteLn(c.FVal); end."#); }
#[test] fn c_override_method() { must_parse(r#"program T; type TBase = class public constructor Create; function Greet: String; end; type TChild = class(TBase) public constructor Create; function Greet: String; end; constructor TBase.Create; begin end; function TBase.Greet: String; begin Result := 'base'; end; constructor TChild.Create; begin inherited Create; end; function TChild.Greet: String; begin Result := 'child'; end; var c: TChild; begin c := TChild.Create; WriteLn(c.Greet()); end."#); }
#[test] fn c_child_own_field() { must_parse(r#"program T; type TBase = class public FX: Integer; constructor Create(X: Integer); end; type TChild = class(TBase) public FY: Integer; constructor Create(X, Y: Integer); function Sum: Integer; end; constructor TBase.Create(X: Integer); begin FX := X; end; constructor TChild.Create(X, Y: Integer); begin inherited Create(X); FY := Y; end; function TChild.Sum: Integer; begin Result := FX + FY; end; var c: TChild; begin c := TChild.Create(10, 20); WriteLn(c.Sum()); end."#); }
#[test] fn c_child_own_method(){ must_parse(r#"program T; type TBase = class public FName: String; constructor Create(N: String); function GetName: String; end; type TChild = class(TBase) public constructor Create(N: String); function Upper: String; end; constructor TBase.Create(N: String); begin FName := N; end; function TBase.GetName: String; begin Result := FName; end; constructor TChild.Create(N: String); begin inherited Create(N); end; function TChild.Upper: String; begin Result := UpperCase(FName); end; var c: TChild; begin c := TChild.Create('hello'); WriteLn(c.GetName()); WriteLn(c.Upper()); end."#); }

// ── Every source from test_new_features.rs ──────────────────────────────────
#[test] fn n_forin()           { must_parse("program T; var item: Integer; begin for item in [10, 20, 30] do WriteLn(item); end."); }
#[test] fn n_forin_str()       { must_parse("program T; var ch: String; begin for ch in 'abc' do WriteLn(ch); end."); }
#[test] fn n_forin_var()       { must_parse("program T; var a: array of Integer; var x: Integer; begin a := [5, 10, 15]; for x in a do WriteLn(x); end."); }
#[test] fn n_compound_add()    { must_parse("program T; var x: Integer; begin x := 10; x += 5; WriteLn(x); end."); }
#[test] fn n_compound_sub()    { must_parse("program T; var x: Integer; begin x := 10; x -= 3; WriteLn(x); end."); }
#[test] fn n_compound_mul()    { must_parse("program T; var x: Integer; begin x := 5; x *= 3; WriteLn(x); end."); }
#[test] fn n_compound_div()    { must_parse("program T; var x: Real; begin x := 10.0; x /= 4.0; WriteLn(x); end."); }
#[test] fn n_compound_str()    { must_parse("program T; var s: String; begin s := 'hello'; s += ' world'; WriteLn(s); end."); }
#[test] fn n_enum_basic()      { must_parse("program T; type TColor = (Red, Green, Blue); var c: TColor; begin c := Green; WriteLn(c); end."); }
#[test] fn n_enum_case()       { must_parse("program T; type TColor = (Red, Green, Blue); var c: TColor; begin c := Blue; case c of 0: WriteLn('red'); 1: WriteLn('green'); 2: WriteLn('blue'); end; end."); }
#[test] fn n_is_check()        { must_parse(r#"program T; type TAnimal = class public FName: String; constructor Create(N: String); end; constructor TAnimal.Create(N: String); begin FName := N; end; var a: TAnimal; begin a := TAnimal.Create('Rex'); if a is TAnimal then WriteLn('yes') else WriteLn('no'); end."#); }
#[test] fn n_as_cast()         { must_parse(r#"program T; type TFoo = class public FVal: Integer; constructor Create(V: Integer); end; constructor TFoo.Create(V: Integer); begin FVal := V; end; var f: TFoo; begin f := TFoo.Create(42); WriteLn((f as TFoo).FVal); end."#); }
#[test] fn n_lambda_basic()    { must_parse(r#"program T; var f: procedure; begin f := procedure begin WriteLn('hello from lambda'); end; f(); end."#); }
#[test] fn n_lambda_params()   { must_parse(r#"program T; var add: function(a, b: Integer): Integer; begin add := function(a, b: Integer): Integer begin Result := a + b; end; WriteLn(add(3, 4)); end."#); }
#[test] fn n_self_field()      { must_parse(r#"program T; type TFoo = class public FVal: Integer; constructor Create(V: Integer); function GetVal: Integer; end; constructor TFoo.Create(V: Integer); begin Self.FVal := V; end; function TFoo.GetVal: Integer; begin Result := Self.FVal; end; var f: TFoo; begin f := TFoo.Create(42); WriteLn(f.GetVal()); end."#); }
#[test] fn n_stringreplace()   { must_parse("program T; begin WriteLn(StringReplace('hello world', 'world', 'pascal')); end."); }
#[test] fn n_stringofchar()    { must_parse("program T; begin WriteLn(StringOfChar('*', 5)); end."); }
#[test] fn n_leftstr()         { must_parse("program T; begin WriteLn(LeftStr('hello', 3)); end."); }
#[test] fn n_typed_const()     { must_parse("program T; const MaxSize: Integer = 100; begin WriteLn(MaxSize); end."); }
#[test] fn n_interface_decl()  { must_parse("program T; type IGreeter = interface function Greet: String; end; begin WriteLn('ok'); end."); }
#[test] fn n_freeandnil()      { must_parse(r#"program T; type TFoo = class public constructor Create; end; constructor TFoo.Create; begin end; var f: TFoo; begin f := TFoo.Create; FreeAndNil(f); if not Assigned(f) then WriteLn('nil') else WriteLn('not nil'); end."#); }

// ── test_programs.rs ────────────────────────────────────────────────────────
#[test] fn p_fizzbuzz()        { must_parse(r#"program T; var i: Integer; begin for i := 1 to 15 do begin if (i mod 15) = 0 then WriteLn('FizzBuzz') else if (i mod 3) = 0 then WriteLn('Fizz') else if (i mod 5) = 0 then WriteLn('Buzz') else WriteLn(i); end; end."#); }
#[test] fn p_sum_100()         { must_parse("program T; var i, s: Integer; begin s := 0; for i := 1 to 100 do s := s + i; WriteLn(s); end."); }
#[test] fn p_gcd()             { must_parse(r#"program T; function GCD(a, b: Integer): Integer; begin if b = 0 then Result := a else Result := GCD(b, a mod b); end; begin WriteLn(GCD(48, 18)); end."#); }
#[test] fn p_is_prime()        { must_parse(r#"program T; function IsPrime(n: Integer): Boolean; var i: Integer; begin if n < 2 then begin Result := false; Exit; end; i := 2; while i * i <= n do begin if (n mod i) = 0 then begin Result := false; Exit; end; i := i + 1; end; Result := true; end; begin WriteLn(IsPrime(2)); WriteLn(IsPrime(4)); WriteLn(IsPrime(7)); WriteLn(IsPrime(9)); WriteLn(IsPrime(13)); end."#); }
#[test] fn p_power()           { must_parse("program T; function Pow(base, exp: Integer): Integer; var i, r: Integer; begin r := 1; for i := 1 to exp do r := r * base; Result := r; end; begin WriteLn(Pow(2, 10)); end."); }
#[test] fn p_reverse_str()     { must_parse("program T; function ReverseStr(s: String): String; var i: Integer; begin Result := ''; for i := Length(s) - 1 downto 0 do Result := Result + s[i]; end; begin WriteLn(ReverseStr('hello')); end."); }
#[test] fn p_count_digits()    { must_parse("program T; function CountDigits(n: Integer): Integer; begin if n < 10 then Result := 1 else Result := 1 + CountDigits(n div 10); end; begin WriteLn(CountDigits(12345)); end."); }
#[test] fn p_sum_digits()      { must_parse("program T; function SumDigits(n: Integer): Integer; begin if n < 10 then Result := n else Result := (n mod 10) + SumDigits(n div 10); end; begin WriteLn(SumDigits(12345)); end."); }
#[test] fn p_max3()            { must_parse("program T; function Max3(a, b, c: Integer): Integer; begin Result := a; if b > Result then Result := b; if c > Result then Result := c; end; begin WriteLn(Max3(3, 7, 5)); WriteLn(Max3(10, 2, 8)); end."); }
#[test] fn p_triangle()        { must_parse("program T; function TriangleArea(base, height: Integer): Real; begin Result := base * height / 2; end; begin WriteLn(TriangleArea(10, 5)); end."); }
#[test] fn p_accumulate()      { must_parse("program T; var i: Integer; var product: Integer; begin product := 1; for i := 1 to 5 do product := product * i; WriteLn(product); end."); }
#[test] fn p_collatz()         { must_parse("program T; function CollatzSteps(n: Integer): Integer; var steps: Integer; begin steps := 0; while n <> 1 do begin if (n mod 2) = 0 then n := n div 2 else n := 3 * n + 1; steps := steps + 1; end; Result := steps; end; begin WriteLn(CollatzSteps(6)); end."); }

// ── test_edge_cases.rs ──────────────────────────────────────────────────────
#[test] fn e_empty()           { must_parse("program T; begin end."); }
#[test] fn e_empty_proc()      { must_parse("program T; procedure P; begin end; begin P; end."); }
#[test] fn e_many_locals()     { must_parse("program T; var a,b,c,d,e: Integer; begin a:=1; b:=2; c:=3; d:=4; e:=5; WriteLn(a+b+c+d+e); end."); }
#[test] fn e_nested_calls()    { must_parse("program T; function F(x: Integer): Integer; begin Result := x + 1; end; begin WriteLn(F(F(F(F(F(0)))))); end."); }
#[test] fn e_zero_iter()       { must_parse("program T; var i: Integer; begin for i := 5 to 3 do WriteLn('x'); WriteLn('done'); end."); }
#[test] fn e_single_char()     { must_parse("program T; begin WriteLn(Length('x')); end."); }
#[test] fn e_shadow()          { must_parse("program T; var x: Integer; procedure Test; var x: Integer; begin x := 99; WriteLn(x); end; begin x := 1; Test; WriteLn(x); end."); }
#[test] fn e_nested_blocks()   { must_parse("program T; var x: Integer; begin x := 1; begin x := 2; WriteLn(x); end; WriteLn(x); end."); }
#[test] fn e_funcs_calling()   { must_parse("program T; function A(x: Integer): Integer; begin Result := x * 2; end; function B(x: Integer): Integer; begin Result := A(x) + 1; end; begin WriteLn(B(5)); end."); }
#[test] fn e_func_in_if()      { must_parse("program T; function IsPositive(x: Integer): Boolean; begin Result := x > 0; end; begin if IsPositive(5) then WriteLn('pos') else WriteLn('neg'); end."); }
#[test] fn e_long_str_build()  { must_parse("program T; var s: String; var i: Integer; begin s := ''; for i := 1 to 5 do s := s + IntToStr(i); WriteLn(s); end."); }
#[test] fn e_mixed_output()    { must_parse("program T; begin WriteLn(1 + 2); WriteLn('hello'); WriteLn(true); WriteLn(3.14); end."); }
#[test] fn e_nested_loop()     { must_parse("program T; var i, j, s: Integer; begin s := 0; for i := 1 to 3 do for j := 1 to 3 do s := s + i * j; WriteLn(s); end."); }
#[test] fn e_while_complex()   { must_parse("program T; var x: Integer; begin x := 100; while (x > 1) and (x > 50) do x := x - 10; WriteLn(x); end."); }
