/// Tests for Char type operations in Object Pascal / Delphi:
/// UpCase, LowCase, character comparison, Ord/Chr patterns,
/// char in conditions, character building and string iteration.
use super::helpers::run_pascal;

// ===================================================================
// UPCASE / LOWERCASE ON CHAR
// ===================================================================

#[test]
fn upcase_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := 'a';
  WriteLn(UpCase(c));
end."#
        ),
        &["A"]
    );
}

#[test]
fn upcase_already_upper() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(UpCase('Z'));
end."#
        ),
        &["Z"]
    );
}

#[test]
fn upcase_non_alpha() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(UpCase('5'));
end."#
        ),
        &["5"]
    );
}

// ===================================================================
// CHAR COMPARISONS
// ===================================================================

#[test]
fn char_less_than() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  if 'a' < 'b' then WriteLn('yes') else WriteLn('no');
end."#
        ),
        &["yes"]
    );
}

#[test]
fn char_greater_than() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  if 'z' > 'a' then WriteLn('yes') else WriteLn('no');
end."#
        ),
        &["yes"]
    );
}

#[test]
fn char_equal() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := 'X';
  if c = 'X' then WriteLn('match') else WriteLn('no match');
end."#
        ),
        &["match"]
    );
}

#[test]
fn char_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := 'A';
  if c <> 'B' then WriteLn('different') else WriteLn('same');
end."#
        ),
        &["different"]
    );
}

// ===================================================================
// ORD AND CHR
// ===================================================================

#[test]
fn ord_of_a() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Ord('A'));
end."#
        ),
        &["65"]
    );
}

#[test]
fn ord_of_zero_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Ord('0'));
end."#
        ),
        &["48"]
    );
}

#[test]
fn chr_to_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Chr(65));
end."#
        ),
        &["A"]
    );
}

#[test]
fn chr_lowercase() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(Chr(97));
end."#
        ),
        &["a"]
    );
}

#[test]
fn ord_chr_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := Chr(Ord('M') + 1);
  WriteLn(c);
end."#
        ),
        &["N"]
    );
}

// ===================================================================
// CHAR IN CASE STATEMENT
// ===================================================================

#[test]
fn char_case_vowel() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := 'a';
  case c of
    'a','e','i','o','u': WriteLn('vowel');
    else WriteLn('consonant');
  end;
end."#
        ),
        &["vowel"]
    );
}

#[test]
fn char_case_consonant() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
begin
  c := 'b';
  case c of
    'a','e','i','o','u': WriteLn('vowel');
    else WriteLn('consonant');
  end;
end."#
        ),
        &["consonant"]
    );
}

// ===================================================================
// CHAR IN STRING BUILDING
// ===================================================================

#[test]
fn char_concat_to_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
    c: Char;
begin
  s := '';
  c := 'H';
  s := s + c;
  c := 'i';
  s := s + c;
  WriteLn(s);
end."#
        ),
        &["Hi"]
    );
}

#[test]
fn count_vowels_in_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
    i, count: Integer;
    c: Char;
begin
  s := 'hello world';
  count := 0;
  for i := 1 to Length(s) do
  begin
    c := s[i];
    if (c = 'a') or (c = 'e') or (c = 'i') or (c = 'o') or (c = 'u') then
      Inc(count);
  end;
  WriteLn(count);
end."#
        ),
        &["3"]
    );
}

#[test]
fn check_digit_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsDigit(c: Char): Boolean;
begin
  Result := (Ord(c) >= Ord('0')) and (Ord(c) <= Ord('9'));
end;
begin
  WriteLn(IsDigit('5'));
  WriteLn(IsDigit('x'));
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn check_alpha_char() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsLetter(c: Char): Boolean;
begin
  Result := ((Ord(c) >= Ord('A')) and (Ord(c) <= Ord('Z'))) or
            ((Ord(c) >= Ord('a')) and (Ord(c) <= Ord('z')));
end;
begin
  WriteLn(IsLetter('G'));
  WriteLn(IsLetter('3'));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// CHAR ARITHMETIC
// ===================================================================

#[test]
fn char_next_letter() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
    i: Integer;
begin
  c := 'A';
  for i := 1 to 5 do
  begin
    Write(c);
    c := Chr(Ord(c) + 1);
  end;
  WriteLn('');
end."#
        ),
        &["ABCDE"]
    );
}

#[test]
fn char_to_digit_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
function DigitValue(c: Char): Integer;
begin
  Result := Ord(c) - Ord('0');
end;
begin
  WriteLn(DigitValue('7'));
  WriteLn(DigitValue('0'));
end."#
        ),
        &["7", "0"]
    );
}

#[test]
fn char_comparison_less_than() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn('a' < 'b'); end."#),
        &["TRUE"]
    );
}

#[test]
fn chr_produces_printable_ascii() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Chr(49)); end."#),
        &["1"]
    );
}

#[test]
fn ord_of_space_is_thirty_two() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Ord(' ')); end."#),
        &["32"]
    );
}

#[test]
fn char_in_set_membership() {
    assert_eq!(
        run_pascal(
            r#"program T; var s: set of Char; begin s := ['a'..'c']; WriteLn('b' in s); end."#
        ),
        &["true"]
    );
}

#[test]
fn pred_char_steps_backward() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(Pred('B')); end."#),
        &["A"]
    );
}
