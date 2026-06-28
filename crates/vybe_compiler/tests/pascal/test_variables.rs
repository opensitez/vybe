use super::helpers::run_pascal;

#[test]
fn var_integer() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 10; WriteLn(x); end."),
        &["10"]
    );
}
#[test]
fn var_string() {
    assert_eq!(
        run_pascal("program T; var s: String; begin s := 'world'; WriteLn(s); end."),
        &["world"]
    );
}
#[test]
fn var_boolean() {
    assert_eq!(
        run_pascal("program T; var b: Boolean; begin b := true; WriteLn(b); end."),
        &["true"]
    );
}
#[test]
fn var_real() {
    assert_eq!(
        run_pascal("program T; var x: Real; begin x := 2.5; WriteLn(x); end."),
        &["2.5"]
    );
}
#[test]
fn var_reassign() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 1; x := 2; x := 3; WriteLn(x); end."),
        &["3"]
    );
}
#[test]
fn var_default_int() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin WriteLn(x); end."),
        &["0"]
    );
}
#[test]
fn var_default_str() {
    assert_eq!(
        run_pascal("program T; var s: String; begin WriteLn(Length(s)); end."),
        &["0"]
    );
}
#[test]
fn var_default_bool() {
    assert_eq!(
        run_pascal("program T; var b: Boolean; begin WriteLn(b); end."),
        &["false"]
    );
}
#[test]
fn var_multiple() {
    assert_eq!(
        run_pascal("program T; var a, b: Integer; begin a := 10; b := 20; WriteLn(a + b); end."),
        &["30"]
    );
}
#[test]
fn var_swap() {
    assert_eq!(
        run_pascal(
            "program T; var a, b, t: Integer; begin a := 1; b := 2; t := a; a := b; b := t; WriteLn(a); WriteLn(b); end."
        ),
        &["2", "1"]
    );
}

// Constants
#[test]
fn const_integer() {
    assert_eq!(
        run_pascal("program T; const N = 42; begin WriteLn(N); end."),
        &["42"]
    );
}
#[test]
fn const_string() {
    assert_eq!(
        run_pascal("program T; const S = 'hello'; begin WriteLn(S); end."),
        &["hello"]
    );
}
#[test]
fn const_bool() {
    assert_eq!(
        run_pascal("program T; const B = true; begin WriteLn(B); end."),
        &["true"]
    );
}
#[test]
fn const_expr() {
    assert_eq!(
        run_pascal("program T; const N = 6 * 7; begin WriteLn(N); end."),
        &["42"]
    );
}
#[test]
fn const_multiple() {
    assert_eq!(
        run_pascal("program T; const A = 10; B = 20; begin WriteLn(A + B); end."),
        &["30"]
    );
}
#[test]
fn const_maxint() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(MaxInt); end."),
        &["2147483647"]
    );
}

#[test]
fn var_shadowing_inner_hides_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Demo;
var x: Integer;
  procedure Inner;
  var x: Integer;
  begin
    x := 2;
    WriteLn(x);
  end;
begin
  x := 1;
  Inner;
  WriteLn(x);
end;
begin
  Demo;
end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn typed_constant_string_immutable() {
    assert_eq!(
        run_pascal(r#"program T; const Greeting: string = 'hi'; begin WriteLn(Greeting); end."#),
        &["hi"]
    );
}

#[test]
fn typed_constant_real_value() {
    assert_eq!(
        run_pascal(r#"program T; const PiApprox: Real = 3.25; begin WriteLn(PiApprox:0:2); end."#),
        &["3.25"]
    );
}

#[test]
fn local_var_initialized_in_declaration() {
    assert_eq!(
        run_pascal(r#"program T; var n: Integer = 7; begin WriteLn(n); end."#),
        &["7"]
    );
}

#[test]
fn global_var_visible_from_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
var g: Integer;
procedure ShowGlobal;
begin
  WriteLn(g);
end;
begin
  g := 99;
  ShowGlobal;
end."#
        ),
        &["99"]
    );
}

#[test]
fn const_expression_used_in_array_bounds() {
    assert_eq!(
        run_pascal(
            r#"program T; const N = 3; var a: array[1..N] of Integer; begin a[2] := 5; WriteLn(a[2]); end."#
        ),
        &["5"]
    );
}

#[test]
fn minint_constant_value() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(MinInt); end."#),
        &["-2147483648"]
    );
}

#[test]
fn multiple_vars_single_var_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T; var a, b, c: Integer; begin a := 1; b := 2; c := 3; WriteLn(a + b + c); end."#
        ),
        &["6"]
    );
}

#[test]
fn byte_var_overflow_wrap_behavior() {
    assert_eq!(
        run_pascal(r#"program T; var b: Byte; begin b := 255; b := b + 1; WriteLn(b); end."#),
        &["0"]
    );
}
