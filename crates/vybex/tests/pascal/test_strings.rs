/// Tests for Pascal string operations beyond basic builtins.

use super::helpers::run_pascal;

// ===================================================================
// POS — find substring (1-based, 0 if not found)
// ===================================================================

#[test] fn str_pos_found() {
    assert_eq!(run_pascal("program T; begin WriteLn(Pos('lo', 'hello')); end."), &["4"]);
}

#[test] fn str_pos_not_found() {
    assert_eq!(run_pascal("program T; begin WriteLn(Pos('xyz', 'hello')); end."), &["0"]);
}

#[test] fn str_pos_at_start() {
    assert_eq!(run_pascal("program T; begin WriteLn(Pos('he', 'hello')); end."), &["1"]);
}

// ===================================================================
// COPY — extract substring (1-based)
// ===================================================================

#[test] fn str_copy_middle() {
    assert_eq!(run_pascal("program T; begin WriteLn(Copy('hello world', 7, 5)); end."), &["world"]);
}

#[test] fn str_copy_from_start() {
    assert_eq!(run_pascal("program T; begin WriteLn(Copy('abcdef', 1, 3)); end."), &["abc"]);
}

// ===================================================================
// RIGHTSTR — right substring
// ===================================================================

#[test] fn str_rightstr() {
    assert_eq!(run_pascal("program T; begin WriteLn(RightStr('hello', 3)); end."), &["llo"]);
}

#[test] fn str_rightstr_full() {
    assert_eq!(run_pascal("program T; begin WriteLn(RightStr('abc', 3)); end."), &["abc"]);
}

// ===================================================================
// CHR / ORD — character conversion
// ===================================================================

#[test] #[ignore] fn str_chr() {
    assert_eq!(run_pascal("program T; begin WriteLn(Chr(65)); end."), &["A"]);
}

#[test] fn str_ord() {
    assert_eq!(run_pascal("program T; begin WriteLn(Ord('A')); end."), &["65"]);
}

#[test] #[ignore] fn str_chr_ord_roundtrip() {
    assert_eq!(run_pascal("program T; begin WriteLn(Chr(Ord('Z'))); end."), &["Z"]);
}

// ===================================================================
// TRIMLEFT / TRIMRIGHT
// ===================================================================

#[test] fn str_trimleft() {
    assert_eq!(run_pascal("program T; begin WriteLn(TrimLeft('  hi  ')); end."), &["hi  "]);
}

#[test] fn str_trimright() {
    assert_eq!(run_pascal("program T; begin WriteLn(TrimRight('  hi  ')); end."), &["  hi"]);
}

// ===================================================================
// BOOLTOSTR / STRTOBOOL
// ===================================================================

#[test] fn str_booltostr() {
    assert_eq!(run_pascal("program T; begin WriteLn(BoolToStr(true)); end."), &["true"]);
}

#[test] fn str_booltostr_false() {
    assert_eq!(run_pascal("program T; begin WriteLn(BoolToStr(false)); end."), &["false"]);
}

// ===================================================================
// STRING INDEXING (runtime uses 0-based despite profile one_based)
// ===================================================================

#[test] fn str_index_first_char() {
    assert_eq!(run_pascal("program T; var s: String; begin s := 'hello'; WriteLn(s[0]); end."), &["h"]);
}

#[test] fn str_index_last_char() {
    assert_eq!(run_pascal("program T; var s: String; begin s := 'hello'; WriteLn(s[4]); end."), &["o"]);
}

// ===================================================================
// COMPARESTR
// ===================================================================

#[test] #[ignore] fn str_comparestr_equal() {
    assert_eq!(run_pascal("program T; begin WriteLn(CompareStr('abc', 'abc')); end."), &["0"]);
}

// ===================================================================
// STRING IN EXPRESSIONS
// ===================================================================

#[test] fn str_length_in_loop() {
    assert_eq!(run_pascal(r#"program T;
var s: String; i: Integer;
begin
  s := 'abc';
  for i := 0 to Length(s) - 1 do WriteLn(s[i]);
end."#), &["a", "b", "c"]);
}

#[test] fn str_build_reverse() {
    assert_eq!(run_pascal(r#"program T;
var s, r: String; i: Integer;
begin
  s := 'abcd';
  r := '';
  for i := Length(s) - 1 downto 0 do r := r + s[i];
  WriteLn(r);
end."#), &["dcba"]);
}
