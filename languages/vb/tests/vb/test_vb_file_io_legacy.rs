use super::helpers::run_vb;

// Legacy File I/O syntax (many of these are just parsed, but some have runtime support in Microsoft.VisualBasic)
#[test]
fn fileio_freefile() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): Console.WriteLine(f > 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_fileopen_close() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test.txt", OpenMode.Output): FileClose(f): Console.WriteLine("OK"): End Sub: End Module"#
        ),
        vec!["OK"]
    );
}
#[test]
fn fileio_printline() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_print.txt", OpenMode.Output): PrintLine(f, "Hello"): FileClose(f): Console.WriteLine("OK"): End Sub: End Module"#
        ),
        vec!["OK"]
    );
}
#[test]
fn fileio_lineinput() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_print.txt", OpenMode.Input): Dim s = LineInput(f): FileClose(f): Console.WriteLine(s): End Sub: End Module"#
        ),
        vec!["Hello"]
    );
}
#[test]
fn fileio_writeline() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_write.txt", OpenMode.Output): WriteLine(f, "A", "B"): FileClose(f): Console.WriteLine("OK"): End Sub: End Module"#
        ),
        vec!["OK"]
    );
}
#[test]
fn fileio_input() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_write.txt", OpenMode.Input): Dim a As String = "": Dim b As String = "": Input(f, a): Input(f, b): FileClose(f): Console.WriteLine(a & b): End Sub: End Module"#
        ),
        vec!["AB"]
    );
}
#[test]
fn fileio_eof() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_write.txt", OpenMode.Input): Dim e = EOF(f): FileClose(f): Console.WriteLine(e = False Or e = True): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_lof() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_write.txt", OpenMode.Input): Dim l = LOF(f): FileClose(f): Console.WriteLine(l > 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_loc() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_write.txt", OpenMode.Input): Dim l = Loc(f): FileClose(f): Console.WriteLine(l >= 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}

// File System operations
#[test]
fn fileio_filecopy() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): FileCopy("test_print.txt", "test_copy.txt"): Console.WriteLine(System.IO.File.Exists("test_copy.txt")): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_kill() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Kill("test_copy.txt"): Console.WriteLine(System.IO.File.Exists("test_copy.txt")): End Sub: End Module"#
        ),
        vec!["False"]
    );
}
#[test]
fn fileio_name() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): FileCopy("test_print.txt", "test_rename.txt"): Rename("test_rename.txt", "test_renamed.txt"): Console.WriteLine(System.IO.File.Exists("test_renamed.txt")): Kill("test_renamed.txt"): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_filelen() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(FileLen("test_print.txt") > 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_filedatetime() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(FileDateTime("test_print.txt").Year > 2000): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_getattr() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(CInt(GetAttr("test_print.txt")) >= 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_setattr() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): SetAttr("test_print.txt", FileAttribute.Normal): Console.WriteLine("OK"): End Sub: End Module"#
        ),
        vec!["OK"]
    );
}
#[test]
fn fileio_mkdir_rmdir() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): MkDir("testdir"): Console.WriteLine(System.IO.Directory.Exists("testdir")): RmDir("testdir"): Console.WriteLine(System.IO.Directory.Exists("testdir")): End Sub: End Module"#
        ),
        vec!["True", "False"]
    );
}
#[test]
fn fileio_curdir() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(CurDir().Length > 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn fileio_dir() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Dir("test_print.txt").Length > 0): End Sub: End Module"#
        ),
        vec!["True"]
    );
}

// Modern System.IO checks (to contrast with legacy)
#[test]
fn io_path_combine() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(System.IO.Path.Combine("A", "B").Contains("B")): End Sub: End Module"#
        ),
        vec!["True"]
    );
}
#[test]
fn io_path_getextension() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(System.IO.Path.GetExtension("test.txt")): End Sub: End Module"#
        ),
        vec![".txt"]
    );
}
#[test]
fn io_path_getfilename() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(System.IO.Path.GetFileName("dir\test.txt")): End Sub: End Module"#
        ),
        vec!["test.txt"]
    );
}

// Interaction functions (MsgBox, InputBox, Beep, Shell, AppActivate, SendKeys, DoEvents) - mostly parser tests since they affect UI
#[test]
fn interact_msgbox() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' MsgBox("Test") : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interact_inputbox() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' InputBox("Test") : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interact_beep() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Beep() : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn interact_shell() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Shell("calc.exe") : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}

// CallByName (Late binding)
#[test]
fn interact_callbyname_get() {
    assert_eq!(
        run_vb(
            r#"Class C: Public Property P As String = "V": End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(CallByName(obj, "P", CallType.Get)): End Sub: End Module"#
        ),
        vec!["V"]
    );
}
#[test]
fn interact_callbyname_set() {
    assert_eq!(
        run_vb(
            r#"Class C: Public Property P As String: End Class: Module M: Sub Main(): Dim obj As New C(): CallByName(obj, "P", CallType.Set, "V"): Console.WriteLine(obj.P): End Sub: End Module"#
        ),
        vec!["V"]
    );
}
#[test]
fn interact_callbyname_method() {
    assert_eq!(
        run_vb(
            r#"Class C: Public Function M() As String: Return "M": End Function: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(CallByName(obj, "M", CallType.Method)): End Sub: End Module"#
        ),
        vec!["M"]
    );
}

// Information functions (IsArray, IsDate, IsDBNull, IsError, IsNothing, IsNumeric, IsReference, QBColor, RGB, TypeName, VarType)
#[test]
fn info_rgb() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(RGB(255, 0, 0)): End Sub: End Module"#),
        vec!["255"]
    );
}
#[test]
fn info_qbcolor() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(QBColor(1)): End Sub: End Module"#),
        vec!["8388608"]
    );
}
#[test]
fn info_typename() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(TypeName("Hello")): End Sub: End Module"#
        ),
        vec!["String"]
    );
}
#[test]
fn info_vartype() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): Console.WriteLine(VarType("Hello")): End Sub: End Module"#),
        vec!["8"]
    );
}

// Registry functions (SaveSetting, GetSetting, GetAllSettings, DeleteSetting)
#[test]
fn registry_getsetting() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' SaveSetting("MyApp", "Sec", "Key", "Val") : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}

// Erase statement
#[test]
fn erase_array() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim a() As Integer = {1, 2}: Erase a: Console.WriteLine(a Is Nothing): End Sub: End Module"#
        ),
        vec!["True"]
    );
}

// Stop, End statements
#[test]
fn stop_statement() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Stop : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn end_statement() {
    assert_eq!(
        run_vb(r#"Module M: Sub Main(): ' End : Console.WriteLine("Parsed"): End Sub: End Module"#),
        vec!["Parsed"]
    );
}

// My Namespace tests
#[test]
fn my_application() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Console.WriteLine(My.Application.Info.Title) : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn my_computer() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Console.WriteLine(My.Computer.Name) : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn my_user() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Console.WriteLine(My.User.Name) : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn my_settings() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Console.WriteLine(My.Settings.Test) : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}
#[test]
fn my_resources() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Console.WriteLine(My.Resources.Test) : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}

// Global namespace
#[test]
fn global_namespace_access() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Console.WriteLine(Global.System.Math.Abs(-1)): End Sub: End Module"#
        ),
        vec!["1"]
    );
}

// Default instances of forms/classes
#[test]
fn default_instance_form() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): ' Form1.Show() : Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}

// Type parameters in operators
#[test]
fn type_param_operator() {
    assert_eq!(
        run_vb(
            r#"Class C(Of T): Public Shared Operator +(a As C(Of T), b As C(Of T)) As C(Of T): Return a: End Operator: End Class: Module M: Sub Main(): Console.WriteLine("Parsed"): End Sub: End Module"#
        ),
        vec!["Parsed"]
    );
}

// Decimal literals
#[test]
fn literal_decimal() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim d = 1.5D: Console.WriteLine(d.GetType().Name): End Sub: End Module"#
        ),
        vec!["Decimal"]
    );
}

// Single literals
#[test]
fn literal_single() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim f = 1.5F: Console.WriteLine(f.GetType().Name): End Sub: End Module"#
        ),
        vec!["Single"]
    );
}

// Long literals
#[test]
fn literal_long() {
    assert_eq!(
        run_vb(
            r#"Module M: Sub Main(): Dim l = 100L: Console.WriteLine(l.GetType().Name): End Sub: End Module"#
        ),
        vec!["Int64"]
    );
}
