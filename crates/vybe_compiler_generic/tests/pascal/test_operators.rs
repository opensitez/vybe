
use super::helpers::run;

// Arithmetic
#[test] fn arith_add()        { assert_eq!(run("program T; begin WriteLn(3 + 4); end."), &["7"]); }
#[test] fn arith_sub()        { assert_eq!(run("program T; begin WriteLn(10 - 3); end."), &["7"]); }
#[test] fn arith_mul()        { assert_eq!(run("program T; begin WriteLn(6 * 7); end."), &["42"]); }
#[test] fn arith_div_real()   { assert_eq!(run("program T; begin WriteLn(10 / 4); end."), &["2.5"]); }
#[test] fn arith_idiv()       { assert_eq!(run("program T; begin WriteLn(10 div 3); end."), &["3"]); }
#[test] fn arith_mod()        { assert_eq!(run("program T; begin WriteLn(10 mod 3); end."), &["1"]); }
#[test] fn arith_neg()        { assert_eq!(run("program T; begin WriteLn(-(3 + 4)); end."), &["-7"]); }
#[test] fn arith_precedence() { assert_eq!(run("program T; begin WriteLn(2 + 3 * 4); end."), &["14"]); }
#[test] fn arith_parens()     { assert_eq!(run("program T; begin WriteLn((2 + 3) * 4); end."), &["20"]); }
#[test] fn arith_chain()      { assert_eq!(run("program T; begin WriteLn(1 + 2 + 3 + 4); end."), &["10"]); }
#[test] fn arith_mixed()      { assert_eq!(run("program T; begin WriteLn(10 - 2 * 3); end."), &["4"]); }
#[test] fn arith_mod_zero()   { assert_eq!(run("program T; begin WriteLn(7 mod 7); end."), &["0"]); }

// Comparison
#[test] fn cmp_eq_true()  { assert_eq!(run("program T; begin if 5 = 5 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_eq_false() { assert_eq!(run("program T; begin if 5 = 6 then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn cmp_ne()       { assert_eq!(run("program T; begin if 5 <> 6 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_lt()       { assert_eq!(run("program T; begin if 3 < 5 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_gt()       { assert_eq!(run("program T; begin if 5 > 3 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_le()       { assert_eq!(run("program T; begin if 3 <= 3 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_ge()       { assert_eq!(run("program T; begin if 5 >= 5 then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_str_eq()   { assert_eq!(run("program T; begin if 'abc' = 'abc' then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn cmp_str_ne()   { assert_eq!(run("program T; begin if 'abc' <> 'xyz' then WriteLn('y') else WriteLn('n'); end."), &["y"]); }

// Boolean / Logical
#[test] fn bool_and_tt()  { assert_eq!(run("program T; begin if true and true then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn bool_and_tf()  { assert_eq!(run("program T; begin if true and false then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn bool_or_tf()   { assert_eq!(run("program T; begin if true or false then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn bool_or_ff()   { assert_eq!(run("program T; begin if false or false then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn bool_not_t()   { assert_eq!(run("program T; begin if not true then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn bool_not_f()   { assert_eq!(run("program T; begin if not false then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn bool_short_and() { assert_eq!(run("program T; begin if false and true then WriteLn('y') else WriteLn('n'); end."), &["n"]); }
#[test] fn bool_short_or()  { assert_eq!(run("program T; begin if true or false then WriteLn('y') else WriteLn('n'); end."), &["y"]); }
#[test] fn bool_compound()  { assert_eq!(run("program T; begin if (3 > 2) and (5 > 4) then WriteLn('y'); end."), &["y"]); }
#[test] fn bool_complex()   { assert_eq!(run("program T; begin if (1 < 2) or (10 < 5) then WriteLn('y'); end."), &["y"]); }
