use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 40: SysUtils & StrUtils Helper Functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_containstext_case_insensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(ContainsText('Hello World', 'WORLD'));
  WriteLn(ContainsText('Hello World', 'PASCAL'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_containsstr_case_sensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(ContainsStr('Hello World', 'World'));
  WriteLn(ContainsStr('Hello World', 'WORLD'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_startstext_case_insensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(StartsText('http', 'HTTP://LOCALHOST'));
  WriteLn(StartsText('ftp', 'HTTP://LOCALHOST'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_startsstr_case_sensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(StartsStr('HTTP', 'HTTP://LOCALHOST'));
  WriteLn(StartsStr('http', 'HTTP://LOCALHOST'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_endstext_case_insensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(EndsText('.TXT', 'document.txt'));
  WriteLn(EndsText('.PNG', 'document.txt'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_endsstr_case_sensitive() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(EndsStr('.txt', 'document.txt'));
  WriteLn(EndsStr('.TXT', 'document.txt'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_leftstr_and_ansileftstr() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(LeftStr('Object Pascal', 6));
end.
"#);
    assert_eq!(out, vec!["Object"]);
}

#[test]
fn test_rightstr_and_ansirightstr() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(RightStr('Object Pascal', 6));
end.
"#);
    assert_eq!(out, vec!["Pascal"]);
}

#[test]
fn test_midstr_and_ansimidstr() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(MidStr('Object Pascal', 8, 6));
end.
"#);
    assert_eq!(out, vec!["Pascal"]);
}

#[test]
fn test_npos_nth_occurrence() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(NPos('a', 'banana', 1));
  WriteLn(NPos('a', 'banana', 2));
  WriteLn(NPos('a', 'banana', 3));
end.
"#);
    assert_eq!(out, vec!["2", "4", "6"]);
}

#[test]
fn test_quotedstr_function() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(QuotedStr('Pascal'));
  WriteLn(QuotedStr('It''s OK'));
end.
"#);
    assert_eq!(out, vec!["'Pascal'", "'It''s OK'"]);
}

#[test]
fn test_dequotedstr_function() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(DequotedStr('''Pascal''', ''''));
end.
"#);
    assert_eq!(out, vec!["Pascal"]);
}

#[test]
fn test_lastdelimiter_function() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var path: String;
begin
  path := '/usr/local/bin/pascal';
  WriteLn(LastDelimiter('/', path));
end.
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_isdelimiter_function() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var text: String;
begin
  text := 'a,b;c';
  WriteLn(IsDelimiter(',;', text, 2));
  WriteLn(IsDelimiter(',;', text, 1));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_wordcount_helper() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(WordCount('One Two Three Four', [' ']));
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_extractword_helper() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(ExtractWord(2, 'One Two Three', [' ']));
end.
"#);
    assert_eq!(out, vec!["Two"]);
}

#[test]
fn test_string_url_scheme_validation() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
function IsSecureUrl(url: String): Boolean;
begin
  Result := StartsText('https://', url);
end;
begin
  WriteLn(IsSecureUrl('HTTPS://EXAMPLE.COM'));
  WriteLn(IsSecureUrl('http://example.com'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_file_extension_check_with_endstext() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
function IsPascalFile(fileName: String): Boolean;
begin
  Result := EndsText('.pas', fileName) or EndsText('.pp', fileName);
end;
begin
  WriteLn(IsPascalFile('main.PAS'));
  WriteLn(IsPascalFile('test.pp'));
  WriteLn(IsPascalFile('doc.txt'));
end.
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_quotedstr_and_dequotedstr_roundtrip() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var orig, quoted, unquoted: String;
begin
  orig := 'SampleString';
  quoted := QuotedStr(orig);
  unquoted := DequotedStr(quoted, '''');
  WriteLn(unquoted = orig);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_leftstr_count_exceeding_length() {
    let out = run_pascal(r#"
program Test;
uses StrUtils;
begin
  WriteLn(LeftStr('Short', 10));
end.
"#);
    assert_eq!(out, vec!["Short"]);
}
