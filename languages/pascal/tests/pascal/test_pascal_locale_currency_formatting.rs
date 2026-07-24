use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 78: Locale & Currency Formatting (TFormatSettings)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_locale_tformatsettings_creation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings;
begin
  fs := TFormatSettings.Create;
  WriteLn(Length(fs.CurrencyString) >= 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_custom_decimal_thousand_separator() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; s: String;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := ',';
  fs.ThousandSeparator := '.';
  s := FormatFloat('#,##0.00', 1234567.89, fs);
  WriteLn(s);
end.
"#);
    assert_eq!(out, vec!["1.234.567,89"]);
}

#[test]
fn test_locale_custom_currency_string() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; s: String;
begin
  fs := TFormatSettings.Create;
  fs.CurrencyString := 'EUR';
  fs.CurrencyFormat := 3; // '123.45 EUR'
  s := FormatCurr('0.00 "', 49.99, fs);
  WriteLn(Pos('49.99', s) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_formatfloat_scientific_notation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FormatFloat('0.00E+00', 1234.56));
end.
"#);
    assert_eq!(out, vec!["1.23E+03"]);
}

#[test]
fn test_locale_formatfloat_padding_zeros() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FormatFloat('00000.00', 42.5));
end.
"#);
    assert_eq!(out, vec!["00042.50"]);
}

#[test]
fn test_locale_formatfloat_optional_digits() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FormatFloat('###.##', 12.3));
end.
"#);
    assert_eq!(out, vec!["12.3"]);
}

#[test]
fn test_locale_strtofloat_with_custom_formatsettings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; val: Double;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := ',';
  val := StrToFloat('123,45', fs);
  WriteLn(val = 123.45);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_strtodate_with_custom_formatsettings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; dt: TDateTime; y, m, d: Word;
begin
  fs := TFormatSettings.Create;
  fs.ShortDateFormat := 'dd/mm/yyyy';
  fs.DateSeparator := '/';
  dt := StrToDate('25/12/2026', fs);
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#);
    assert_eq!(out, vec!["2026-12-25"]);
}

#[test]
fn test_locale_shortdateformat_customization() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; dt: TDateTime;
begin
  fs := TFormatSettings.Create;
  fs.ShortDateFormat := 'yyyy.mm.dd';
  dt := EncodeDate(2026, 9, 1);
  WriteLn(DateToStr(dt, fs));
end.
"#);
    assert_eq!(out, vec!["2026.09.01"]);
}

#[test]
fn test_locale_shorttimeformat_customization() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; dt: TDateTime;
begin
  fs := TFormatSettings.Create;
  fs.ShortTimeFormat := 'hh-nn-ss';
  dt := EncodeTime(15, 45, 30, 0);
  WriteLn(TimeToStr(dt, fs));
end.
"#);
    assert_eq!(out, vec!["15-45-30"]);
}

#[test]
fn test_locale_booltostr_with_custom_boolstrs() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  SetVerifyBoolStrs(False);
  WriteLn(BoolToStr(True, True));
  WriteLn(BoolToStr(False, True));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_locale_currency_negative_format() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; s: String;
begin
  fs := TFormatSettings.Create;
  fs.CurrencyString := '$';
  s := FormatCurr('$#,##0.00;($#,##0.00)', -50.0, fs);
  WriteLn(Pos('50.00', s) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_strtocurr_with_formatsettings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; c: Currency;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := '.';
  c := StrToCurr('199.99', fs);
  WriteLn(c = 199.99);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_formatfloat_percentage() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(FormatFloat('0.0%', 0.25));
end.
"#);
    assert_eq!(out, vec!["25.0%"]);
}

#[test]
fn test_locale_month_names_access() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings;
begin
  fs := TFormatSettings.Create;
  WriteLn(Length(fs.LongMonthNames[1]) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_day_names_access() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings;
begin
  fs := TFormatSettings.Create;
  WriteLn(Length(fs.LongDayNames[1]) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_locale_formatdatetime_with_formatsettings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; dt: TDateTime;
begin
  fs := TFormatSettings.Create;
  dt := EncodeDate(2026, 11, 30);
  WriteLn(FormatDateTime('yyyy/mm/dd', dt, fs));
end.
"#);
    assert_eq!(out, vec!["2026/11/30"]);
}

#[test]
fn test_locale_currtostr_with_formatsettings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings; c: Currency;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := ',';
  c := 75.50;
  WriteLn(CurrToStr(c, fs));
end.
"#);
    assert_eq!(out, vec!["75,5"]);
}

#[test]
fn test_locale_formatfloat_comma_groups() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings;
begin
  fs := TFormatSettings.Create;
  fs.ThousandSeparator := ',';
  fs.DecimalSeparator := '.';
  WriteLn(FormatFloat('#,##0', 1000000, fs));
end.
"#);
    assert_eq!(out, vec!["1,000,000"]);
}

#[test]
fn test_locale_invariant_format_settings() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fs: TFormatSettings;
begin
  fs := TFormatSettings.Invariant;
  WriteLn(fs.DecimalSeparator = '.');
  WriteLn(fs.ThousandSeparator = ',');
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}
