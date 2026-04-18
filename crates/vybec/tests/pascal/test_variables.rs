use super::helpers;
use helpers::run;

#[test] fn var_integer()      { assert_eq!(run("program T; var x: Integer; begin x := 10; WriteLn(x); end."), &["10"]); }
#[test] fn var_string()       { assert_eq!(run("program T; var s: String; begin s := 'world'; WriteLn(s); end."), &["world"]); }
#[test] fn var_boolean()      { assert_eq!(run("program T; var b: Boolean; begin b := true; WriteLn(b); end."), &["true"]); }
#[test] fn var_real()         { assert_eq!(run("program T; var x: Real; begin x := 2.5; WriteLn(x); end."), &["2.5"]); }
#[test] fn var_reassign()     { assert_eq!(run("program T; var x: Integer; begin x := 1; x := 2; x := 3; WriteLn(x); end."), &["3"]); }
#[test] fn var_default_int()  { assert_eq!(run("program T; var x: Integer; begin WriteLn(x); end."), &["0"]); }
#[test] fn var_default_str()  { assert_eq!(run("program T; var s: String; begin WriteLn(Length(s)); end."), &["0"]); }
#[test] fn var_default_bool() { assert_eq!(run("program T; var b: Boolean; begin WriteLn(b); end."), &["false"]); }
#[test] fn var_multiple()     { assert_eq!(run("program T; var a, b: Integer; begin a := 10; b := 20; WriteLn(a + b); end."), &["30"]); }
#[test] fn var_swap() {
    assert_eq!(run("program T; var a, b, t: Integer; begin a := 1; b := 2; t := a; a := b; b := t; WriteLn(a); WriteLn(b); end."), &["2", "1"]);
}

// Constants
#[test] fn const_integer()    { assert_eq!(run("program T; const N = 42; begin WriteLn(N); end."), &["42"]); }
#[test] fn const_string()     { assert_eq!(run("program T; const S = 'hello'; begin WriteLn(S); end."), &["hello"]); }
#[test] fn const_bool()       { assert_eq!(run("program T; const B = true; begin WriteLn(B); end."), &["true"]); }
#[test] fn const_expr()       { assert_eq!(run("program T; const N = 6 * 7; begin WriteLn(N); end."), &["42"]); }
#[test] fn const_multiple()   { assert_eq!(run("program T; const A = 10; B = 20; begin WriteLn(A + B); end."), &["30"]); }
#[test] fn const_maxint()     { assert_eq!(run("program T; begin WriteLn(MaxInt); end."), &["2147483647"]); }
