use super::helpers;
use helpers::run;

// If/Then/Else
#[test] fn if_true()       { assert_eq!(run("program T; begin if true then WriteLn('y'); end."), &["y"]); }
#[test] fn if_false()      { assert_eq!(run("program T; begin if false then WriteLn('y'); end."), &[] as &[&str]); }
#[test] fn if_else_true()  { assert_eq!(run("program T; begin if true then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn if_else_false() { assert_eq!(run("program T; begin if false then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn if_comparison() { assert_eq!(run("program T; var x: Integer; begin x := 5; if x > 3 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn if_nested() {
    assert_eq!(run(r#"program T; var x: Integer; begin x := 10;
      if x > 5 then if x > 8 then WriteLn('big') else WriteLn('med') else WriteLn('small'); end."#), &["big"]);
}
#[test] fn if_chained() {
    assert_eq!(run(r#"program T; var x: Integer; begin x := 2;
      if x = 1 then WriteLn('one')
      else if x = 2 then WriteLn('two')
      else WriteLn('other'); end."#), &["two"]);
}
#[test] fn if_block() {
    assert_eq!(run("program T; begin if true then begin WriteLn('a'); WriteLn('b'); end; end."), &["a", "b"]);
}

// For loops
#[test] fn for_up()     { assert_eq!(run("program T; var i: Integer; begin for i := 1 to 5 do WriteLn(i); end."), &["1","2","3","4","5"]); }
#[test] fn for_down()   { assert_eq!(run("program T; var i: Integer; begin for i := 5 downto 1 do WriteLn(i); end."), &["5","4","3","2","1"]); }
#[test] fn for_single() { assert_eq!(run("program T; var i: Integer; begin for i := 1 to 1 do WriteLn(i); end."), &["1"]); }
#[test] fn for_block() {
    assert_eq!(run("program T; var i, s: Integer; begin s := 0; for i := 1 to 5 do begin s := s + i; end; WriteLn(s); end."), &["15"]);
}
#[test] fn for_nested() {
    assert_eq!(run("program T; var i, j: Integer; begin for i := 1 to 2 do for j := 1 to 2 do WriteLn(i * 10 + j); end."), &["11","12","21","22"]);
}
#[test] fn for_zero_iterations() {
    assert_eq!(run("program T; var i: Integer; begin for i := 5 to 3 do WriteLn('x'); WriteLn('done'); end."), &["done"]);
}

// While loops
#[test] fn while_basic() {
    assert_eq!(run("program T; var i: Integer; begin i := 0; while i < 3 do begin WriteLn(i); i := i + 1; end; end."), &["0","1","2"]);
}
#[test] fn while_false() {
    assert_eq!(run("program T; begin while false do WriteLn('x'); end."), &[] as &[&str]);
}
#[test] fn while_countdown() {
    assert_eq!(run("program T; var i: Integer; begin i := 3; while i > 0 do begin WriteLn(i); i := i - 1; end; end."), &["3","2","1"]);
}

// Repeat/Until
#[test] fn repeat_basic() {
    assert_eq!(run("program T; var i: Integer; begin i := 1; repeat WriteLn(i); i := i + 1; until i > 3; end."), &["1","2","3"]);
}
#[test] fn repeat_once() {
    assert_eq!(run("program T; begin repeat WriteLn('once'); until true; end."), &["once"]);
}

// Break
#[test] fn break_for() {
    assert_eq!(run("program T; var i: Integer; begin for i := 1 to 10 do begin if i > 3 then Break; WriteLn(i); end; end."), &["1","2","3"]);
}
#[test] fn break_while() {
    assert_eq!(run("program T; var i: Integer; begin i := 0; while true do begin i := i + 1; if i > 2 then Break; WriteLn(i); end; end."), &["1","2"]);
}

// Case
#[test] fn case_basic() {
    assert_eq!(run("program T; var x: Integer; begin x := 2; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."), &["two"]);
}
#[test] fn case_else() {
    assert_eq!(run("program T; var x: Integer; begin x := 5; case x of 1: WriteLn('one'); 2: WriteLn('two'); else WriteLn('other'); end; end."), &["other"]);
}
#[test] fn case_first() {
    assert_eq!(run("program T; var x: Integer; begin x := 1; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."), &["one"]);
}
#[test] fn case_last() {
    assert_eq!(run("program T; var x: Integer; begin x := 3; case x of 1: WriteLn('one'); 2: WriteLn('two'); 3: WriteLn('three'); end; end."), &["three"]);
}
