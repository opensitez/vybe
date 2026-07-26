use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 7: Default Parameter Values in Routines & Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_default_param_single_integer() {
    let out = run_pascal(
        r#"
program Test;
function AddBonus(base: Integer; bonus: Integer = 10): Integer;
begin
  Result := base + bonus;
end;
begin
  WriteLn(AddBonus(50));
  WriteLn(AddBonus(50, 25));
end.
"#,
    );
    assert_eq!(out, vec!["60", "75"]);
}

#[test]
fn test_default_param_single_string() {
    let out = run_pascal(
        r#"
program Test;
function Greet(name: String; prefix: String = 'Hello'): String;
begin
  Result := prefix + ' ' + name;
end;
begin
  WriteLn(Greet('Alice'));
  WriteLn(Greet('Bob', 'Welcome'));
end.
"#,
    );
    assert_eq!(out, vec!["Hello Alice", "Welcome Bob"]);
}

#[test]
fn test_default_param_boolean_flag() {
    let out = run_pascal(
        r#"
program Test;
procedure LogMsg(msg: String; verbose: Boolean = False);
begin
  if verbose then WriteLn('[VERBOSE] ' + msg)
  else WriteLn('[INFO] ' + msg);
end;
begin
  LogMsg('System started');
  LogMsg('Debug details', True);
end.
"#,
    );
    assert_eq!(
        out,
        vec!["[INFO] System started", "[VERBOSE] Debug details"]
    );
}

#[test]
fn test_default_param_multiple_defaults() {
    let out = run_pascal(
        r#"
program Test;
function CalcBoxVolume(w: Integer = 1; h: Integer = 1; d: Integer = 1): Integer;
begin
  Result := w * h * d;
end;
begin
  WriteLn(CalcBoxVolume);
  WriteLn(CalcBoxVolume(5));
  WriteLn(CalcBoxVolume(5, 4));
  WriteLn(CalcBoxVolume(5, 4, 3));
end.
"#,
    );
    assert_eq!(out, vec!["1", "5", "20", "60"]);
}

#[test]
fn test_default_param_enum_type() {
    let out = run_pascal(
        r#"
program Test;
type TAlign = (alLeft, alCenter, alRight);
function FormatAlign(text: String; align: TAlign = alCenter): String;
begin
  Result := Ord(align).ToString + ':' + text;
end;
begin
  WriteLn(FormatAlign('Title'));
  WriteLn(FormatAlign('Body', alLeft));
end.
"#,
    );
    assert_eq!(out, vec!["1:Title", "0:Body"]);
}

#[test]
fn test_default_param_floating_point() {
    let out = run_pascal(
        r#"
program Test;
function MultiplyFactor(val: Real; factor: Real = 1.5): Real;
begin
  Result := val * factor;
end;
begin
  WriteLn(MultiplyFactor(10.0));
  WriteLn(MultiplyFactor(10.0, 2.0));
end.
"#,
    );
    assert_eq!(out, vec!["15", "20"]);
}

#[test]
fn test_default_param_in_class_method() {
    let out = run_pascal(
        r#"
program Test;
type TPrinter = class
  public procedure PrintHeader(title: String = 'DEFAULT TITLE');
end;
procedure TPrinter.PrintHeader(title: String);
begin
  WriteLn('=== ' + title + ' ===');
end;
var p: TPrinter;
begin
  p := TPrinter.Create;
  p.PrintHeader;
  p.PrintHeader('CUSTOM TITLE');
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["=== DEFAULT TITLE ===", "=== CUSTOM TITLE ==="]);
}

#[test]
fn test_default_param_in_constructor() {
    let out = run_pascal(
        r#"
program Test;
type TConfig = class
  public Timeout: Integer;
  constructor Create(ATimeout: Integer = 30);
end;
constructor TConfig.Create(ATimeout: Integer);
begin
  Timeout := ATimeout;
end;
var c1, c2: TConfig;
begin
  c1 := TConfig.Create;
  c2 := TConfig.Create(60);
  WriteLn(c1.Timeout);
  WriteLn(c2.Timeout);
  c1.Free; c2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["30", "60"]);
}

#[test]
fn test_default_param_char_literal() {
    let out = run_pascal(
        r#"
program Test;
function PadString(s: String; len: Integer; ch: Char = ' '): String;
begin
  Result := s;
  while Length(Result) < len do
    Result := Result + ch;
end;
begin
  WriteLn(PadString('Hi', 5));
  WriteLn(PadString('Hi', 5, '*'));
end.
"#,
    );
    assert_eq!(out, vec!["Hi   ", "Hi***"]);
}

#[test]
fn test_default_param_constant_expression() {
    let out = run_pascal(
        r#"
program Test;
const DefaultBase = 100;
function BaseOffset(val: Integer; offset: Integer = DefaultBase + 20): Integer;
begin
  Result := val + offset;
end;
begin
  WriteLn(BaseOffset(5));
end.
"#,
    );
    assert_eq!(out, vec!["125"]);
}

#[test]
fn test_default_param_nil_pointer() {
    let out = run_pascal(
        r#"
program Test;
procedure ProcessPointer(ptr: PInteger = nil);
begin
  if ptr = nil then WriteLn('NIL')
  else WriteLn(ptr^);
end;
var val: Integer;
begin
  ProcessPointer;
  val := 42;
  ProcessPointer(@val);
end.
"#,
    );
    assert_eq!(out, vec!["NIL", "42"]);
}

#[test]
fn test_default_param_overloaded_routines() {
    let out = run_pascal(
        r#"
program Test;
procedure OutputVal(v: Integer; prefix: String = 'INT: '); overload;
begin
  WriteLn(prefix + v.ToString);
end;
procedure OutputVal(v: String; prefix: String = 'STR: '); overload;
begin
  WriteLn(prefix + v);
end;
begin
  OutputVal(42);
  OutputVal('hello');
end.
"#,
    );
    assert_eq!(out, vec!["INT: 42", "STR: hello"]);
}

#[test]
fn test_default_param_with_const_modifier() {
    let out = run_pascal(
        r#"
program Test;
function AppendSuffix(const s: String; const suffix: String = '.txt'): String;
begin
  Result := s + suffix;
end;
begin
  WriteLn(AppendSuffix('document'));
  WriteLn(AppendSuffix('image', '.png'));
end.
"#,
    );
    assert_eq!(out, vec!["document.txt", "image.png"]);
}

#[test]
fn test_default_param_mixed_explicit_and_defaults() {
    let out = run_pascal(
        r#"
program Test;
function ComputeTax(amount: Real; rate: Real = 0.05; shipping: Real = 10.0): Real;
begin
  Result := amount + (amount * rate) + shipping;
end;
begin
  WriteLn(ComputeTax(100.0));
  WriteLn(ComputeTax(100.0, 0.10));
  WriteLn(ComputeTax(100.0, 0.10, 5.0));
end.
"#,
    );
    assert_eq!(out, vec!["115", "115", "115"]);
}

#[test]
fn test_default_param_in_nested_procedure() {
    let out = run_pascal(
        r#"
program Test;
procedure MainProc;
  procedure InnerProc(step: Integer = 1);
  begin
    WriteLn('Step ' + step.ToString);
  end;
begin
  InnerProc;
  InnerProc(5);
end;
begin
  MainProc;
end.
"#,
    );
    assert_eq!(out, vec!["Step 1", "Step 5"]);
}

#[test]
fn test_default_param_subrange_type() {
    let out = run_pascal(
        r#"
program Test;
type TLevel = 1..5;
procedure SetLevel(lvl: TLevel = 1);
begin
  WriteLn(lvl);
end;
begin
  SetLevel;
  SetLevel(3);
end.
"#,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn test_default_param_hexadecimal_value() {
    let out = run_pascal(
        r#"
program Test;
procedure MaskBits(val: Integer; mask: Integer = $FF);
begin
  WriteLn(val and mask);
end;
begin
  MaskBits($1234);
  MaskBits($1234, $0F);
end.
"#,
    );
    assert_eq!(out, vec!["52", "4"]);
}

#[test]
fn test_default_param_empty_string() {
    let out = run_pascal(
        r#"
program Test;
function WrapString(s: String; prefix: String = ''; suffix: String = ''): String;
begin
  Result := prefix + s + suffix;
end;
begin
  WriteLn(WrapString('Core'));
  WriteLn(WrapString('Core', '<', '>'));
end.
"#,
    );
    assert_eq!(out, vec!["Core", "<Core>"]);
}

#[test]
fn test_default_param_negative_integer() {
    let out = run_pascal(
        r#"
program Test;
function AdjustScore(score: Integer; penalty: Integer = -5): Integer;
begin
  Result := score + penalty;
end;
begin
  WriteLn(AdjustScore(100));
  WriteLn(AdjustScore(100, -10));
end.
"#,
    );
    assert_eq!(out, vec!["95", "90"]);
}

#[test]
fn test_default_param_class_static_method() {
    let out = run_pascal(
        r#"
program Test;
type TMathUtils = class
  public class function Multiply(a: Integer; b: Integer = 2): Integer;
end;
class function TMathUtils.Multiply(a: Integer; b: Integer): Integer;
begin
  Result := a * b;
end;
begin
  WriteLn(TMathUtils.Multiply(10));
  WriteLn(TMathUtils.Multiply(10, 5));
end.
"#,
    );
    assert_eq!(out, vec!["20", "50"]);
}
