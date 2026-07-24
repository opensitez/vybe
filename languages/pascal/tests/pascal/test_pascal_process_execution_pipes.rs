use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 79: Process Execution, Pipes & System Execution
// ═══════════════════════════════════════════════════════════

#[test]
fn test_process_tprocess_instantiation() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  WriteLn(proc <> nil);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_tprocess_executable_and_options() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/bin/echo';
  proc.Parameters.Add('HelloProcess');
  proc.Options := [poWaitOnExit, poUsePipes];
  WriteLn(proc.Executable);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["/bin/echo"]);
}

#[test]
fn test_process_execute_echo_stdout() {
    let out = run_pascal(r#"
program Test;
uses Process, Classes;
var proc: TProcess; sl: TStringList;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/bin/echo';
  proc.Parameters.Add('ProcessPipeOutput');
  proc.Options := [poWaitOnExit, poUsePipes];
  proc.Execute;

  sl := TStringList.Create;
  sl.LoadFromStream(proc.Output);
  WriteLn(Trim(sl.Text));

  sl.Free; proc.Free;
end.
"#);
    assert_eq!(out, vec!["ProcessPipeOutput"]);
}

#[test]
fn test_process_exitcode_success() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/usr/bin/true';
  proc.Options := [poWaitOnExit];
  proc.Execute;
  WriteLn(proc.ExitCode = 0);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_exitcode_failure() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/usr/bin/false';
  proc.Options := [poWaitOnExit];
  proc.Execute;
  WriteLn(proc.ExitCode <> 0);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_multiple_parameters() {
    let out = run_pascal(r#"
program Test;
uses Process, Classes;
var proc: TProcess; sl: TStringList;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/bin/echo';
  proc.Parameters.Add('Param1');
  proc.Parameters.Add('Param2');
  proc.Options := [poWaitOnExit, poUsePipes];
  proc.Execute;

  sl := TStringList.Create;
  sl.LoadFromStream(proc.Output);
  WriteLn(Trim(sl.Text));

  sl.Free; proc.Free;
end.
"#);
    assert_eq!(out, vec!["Param1 Param2"]);
}

#[test]
fn test_process_current_directory_setting() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.CurrentDirectory := '/tmp';
  WriteLn(proc.CurrentDirectory);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["/tmp"]);
}

#[test]
fn test_process_running_check() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  WriteLn(proc.Running);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_process_environment_strings() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Environment.Add('MY_PROC_ENV=123');
  WriteLn(proc.Environment.Count > 0);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_executeprocess_routine() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var res: Integer;
begin
  res := ExecuteProcess('/usr/bin/true', []);
  WriteLn(res = 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_stderr_redirection() {
    let out = run_pascal(r#"
program Test;
uses Process, Classes;
var proc: TProcess; sl: TStringList;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/bin/sh';
  proc.Parameters.Add('-c');
  proc.Parameters.Add('echo ErrorOnStderr >&2');
  proc.Options := [poWaitOnExit, poUsePipes];
  proc.Execute;

  sl := TStringList.Create;
  sl.LoadFromStream(proc.Stderr);
  WriteLn(Trim(sl.Text));

  sl.Free; proc.Free;
end.
"#);
    assert_eq!(out, vec!["ErrorOnStderr"]);
}

#[test]
fn test_process_postderrtooutput_option() {
    let out = run_pascal(r#"
program Test;
uses Process, Classes;
var proc: TProcess; sl: TStringList;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/bin/sh';
  proc.Parameters.Add('-c');
  proc.Parameters.Add('echo MergedError >&2');
  proc.Options := [poWaitOnExit, poUsePipes, poStderrToOutPut];
  proc.Execute;

  sl := TStringList.Create;
  sl.LoadFromStream(proc.Output);
  WriteLn(Trim(sl.Text));

  sl.Free; proc.Free;
end.
"#);
    assert_eq!(out, vec!["MergedError"]);
}

#[test]
fn test_process_stdin_writing() {
    let out = run_pascal(r#"
program Test;
uses Process, Classes;
var proc: TProcess; inputStr: String; sl: TStringList;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/usr/bin/tr';
  proc.Parameters.Add('a-z');
  proc.Parameters.Add('A-Z');
  proc.Options := [poUsePipes];
  proc.Execute;

  inputStr := 'hello stdin' + #10;
  proc.Input.WriteBuffer(inputStr[1], Length(inputStr));
  proc.CloseInput;

  proc.WaitOnExit;

  sl := TStringList.Create;
  sl.LoadFromStream(proc.Output);
  WriteLn(Trim(sl.Text));

  sl.Free; proc.Free;
end.
"#);
    assert_eq!(out, vec!["HELLO STDIN"]);
}

#[test]
fn test_process_protection_finally() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  try
    WriteLn('ProcessCreated');
  finally
    proc.Free;
    WriteLn('ProcessFreedInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["ProcessCreated", "ProcessFreedInFinally"]);
}

#[test]
fn test_process_startupoptions() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.StartupOptions := [suoUseShowWindow];
  proc.ShowWindow := swoHide;
  WriteLn(proc.ShowWindow = swoHide);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_priority_setting() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Priority := ppNormal;
  WriteLn(proc.Priority = ppNormal);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_pipe_buffer_size() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.PipeBufferSize := 8192;
  WriteLn(proc.PipeBufferSize = 8192);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_non_existent_executable_raises_exception() {
    let out = run_pascal(r#"
program Test;
uses Process, SysUtils;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.Executable := '/path/to/non_existent_binary_xyz';
  try
    proc.Execute;
    proc.WaitOnExit;
  except
    on E: Exception do WriteLn('ExecFailedException');
  end;
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["ExecFailedException"]);
}

#[test]
fn test_process_active_process_handle() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  WriteLn(proc.Handle = 0);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_process_custom_commandline_prop() {
    let out = run_pascal(r#"
program Test;
uses Process;
var proc: TProcess;
begin
  proc := TProcess.Create(nil);
  proc.CommandLine := '/bin/echo CustomCommandLine';
  WriteLn(Length(proc.CommandLine) > 0);
  proc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}
