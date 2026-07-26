use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 74: Environment Variables & Command Line Arguments
// ═══════════════════════════════════════════════════════════

#[test]
fn test_env_paramcount_gte_zero() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(ParamCount >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_paramstr_zero_executable_path() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Length(ParamStr(0)) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_getenvironmentvariable_path() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var pathVal: String;
begin
  pathVal := GetEnvironmentVariable('PATH');
  WriteLn(Length(pathVal) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_setenvironmentvariable_custom() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  SetEnvironmentVariable('TEST_ENV_VAR', 'PascalVal');
  WriteLn(GetEnvironmentVariable('TEST_ENV_VAR'));
end.
"#,
    );
    assert_eq!(out, vec!["PascalVal"]);
}

#[test]
fn test_env_findcmdlineswitch_missing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var value: String;
begin
  WriteLn(FindCmdLineSwitch('missing_flag'));
  WriteLn(FindCmdLineSwitch('missing_switch', value));
end.
"#,
    );
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn test_env_exitcode_variable() {
    let out = run_pascal(
        r#"
program Test;
begin
  ExitCode := 0;
  WriteLn(ExitCode);
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_env_getenvironmentvariable_nonexistent() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(GetEnvironmentVariable('NON_EXISTENT_VAR_XYZ_99')) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_getcmdline() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(CmdLine) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_paramstr_out_of_bounds() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Length(ParamStr(9999)) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_getenv_system_function() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(Length(GetEnv('PATH')) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_setenv_overwrite() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  SetEnvironmentVariable('MY_VAR', 'Val1');
  SetEnvironmentVariable('MY_VAR', 'Val2');
  WriteLn(GetEnvironmentVariable('MY_VAR'));
end.
"#,
    );
    assert_eq!(out, vec!["Val2"]);
}

#[test]
fn test_env_setenv_empty_clears() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  SetEnvironmentVariable('TEMP_VAR', 'ToClear');
  SetEnvironmentVariable('TEMP_VAR', '');
  WriteLn(Length(GetEnvironmentVariable('TEMP_VAR')) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_paramcount_loop_safety() {
    let out = run_pascal(
        r#"
program Test;
var i: Integer;
begin
  for i := 1 to ParamCount do
  begin
    WriteLn(ParamStr(i));
  end;
  WriteLn('ParamLoopFinished');
end.
"#,
    );
    assert_eq!(out, vec!["ParamLoopFinished"]);
}

#[test]
fn test_env_custom_cmdline_parse_helper() {
    let out = run_pascal(
        r#"
program Test;
function ParseIntArg(const argStr: String; defaultVal: Integer): Integer;
begin
  if argStr = '' then Result := defaultVal
  else Result := StrToIntDef(argStr, defaultVal);
end;
begin
  WriteLn(ParseIntArg('', 100));
  WriteLn(ParseIntArg('250', 100));
end.
"#,
    );
    assert_eq!(out, vec!["100", "250"]);
}

#[test]
fn test_env_osversion_platform() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(TOSVersion.ToString) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_tosversion_architecture() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Ord(TOSVersion.Architecture) >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_tosversion_platform_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(Ord(TOSVersion.Platform) >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_tosversion_major_minor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(TOSVersion.Major >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_env_environment_variable_case_insensitivity() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  SetEnvironmentVariable('CASING_VAR', 'CaseTest');
  WriteLn(GetEnvironmentVariable('casing_var'));
end.
"#,
    );
    assert_eq!(out, vec!["CaseTest"]);
}

#[test]
fn test_env_exitcode_custom_status() {
    let out = run_pascal(
        r#"
program Test;
procedure FailWithStatus(code: Integer);
begin
  ExitCode := code;
end;
begin
  FailWithStatus(42);
  WriteLn(ExitCode);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}
