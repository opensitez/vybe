//! `[CallerMemberName]`, `[CallerLineNumber]`, and `[CallerFilePath]` on optional parameters.

csharp_cases! {
    caller_member_name_from_instance_method => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
class App { public void Run() { Trace.Show(); } }
new App().Run();"#,
        ["Run"]
    };

    caller_member_name_from_static_method => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
    public static void Ping() { Show(); }
}
Trace.Ping();"#,
        ["Ping"]
    };

    caller_member_name_from_property_getter => {
        r#"class Box {
    int _v = 5;
    public int Value {
        get {
            Report();
            return _v;
        }
    }
    void Report([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Console.WriteLine(new Box().Value);"#,
        ["Value", "5"]
    };

    caller_member_name_from_property_setter => {
        r#"class Box {
    int _v;
    public int Value {
        set {
            Report();
            _v = value;
        }
        get => _v;
    }
    void Report([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
var b = new Box(); b.Value = 9; Console.WriteLine(b.Value);"#,
        ["Value", "9"]
    };

    caller_member_name_from_constructor => {
        r#"class Node {
    public Node() { Trace(); }
    void Trace([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Node();"#,
        [".ctor"]
    };

    caller_member_name_from_indexer_getter => {
        r#"class Row {
    int[] cells = { 1, 2, 3 };
    public int this[int i] {
        get {
            LogAccess();
            return cells[i];
        }
    }
    void LogAccess([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Console.WriteLine(new Row()[1]);"#,
        ["Item", "2"]
    };

    caller_member_name_from_indexer_setter => {
        r#"class Row {
    int[] cells = new int[3];
    public int this[int i] {
        set {
            LogWrite();
            cells[i] = value;
        }
    }
    void LogWrite([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
var r = new Row(); r[0] = 7; Console.WriteLine(r[0]);"#,
        ["Item", "7"]
    };

    caller_member_name_from_operator_method => {
        r#"class Num {
    public int V;
    public static Num operator +(Num a, Num b) {
        Log();
        return new Num { V = a.V + b.V };
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Console.WriteLine((new Num { V = 1 } + new Num { V = 2 }).V);"#,
        ["op_Addition", "3"]
    };

    caller_member_name_from_nested_type_method => {
        r#"class Outer {
    public class Inner { public void Work() { Trace.Show(); } }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Outer.Inner().Work();"#,
        ["Work"]
    };

    caller_member_name_from_event_handler_style => {
        r#"class Btn {
    public void Click() { OnClick(); }
    void OnClick([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Btn().Click();"#,
        ["Click"]
    };

    caller_member_name_explicit_argument_overrides_default => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
class App { public void Run() { Trace.Show("manual"); } }
new App().Run();"#,
        ["manual"]
    };

    caller_member_name_from_local_function => {
        r#"class App {
    public void Run() {
        Local();
        void Local() {
            Trace.Show();
        }
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new App().Run();"#,
        ["Local"]
    };

    caller_line_number_from_single_call_site => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
class App {
    public void Run() {
        Trace.Show();
    }
}
new App().Run();"#,
        ["6"]
    };

    caller_line_number_two_calls_different_lines => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
Trace.Show();
Trace.Show();"#,
        ["4", "5"]
    };

    caller_line_number_from_property_getter => {
        r#"class Box {
    int _v = 1;
    public int Value {
        get {
            Trace.Show();
            return _v;
        }
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
Console.WriteLine(new Box().Value);"#,
        ["5", "1"]
    };

    caller_line_number_explicit_value_overrides_injection => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
Trace.Show(99);"#,
        ["99"]
    };

    caller_line_number_with_other_optional_params => {
        r#"class Trace {
    public static void Show(string tag, [System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(tag + ":" + line);
}
Trace.Show("mark");"#,
        ["mark:4"]
    };

    caller_file_path_is_non_empty_when_omitted => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerFilePath] string path = "") => Console.WriteLine(path.Length > 0);
}
Trace.Show();"#,
        ["True"]
    };

    caller_file_path_explicit_argument_used => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerFilePath] string path = "") => Console.WriteLine(path);
}
Trace.Show("/tmp/sample.cs");"#,
        ["/tmp/sample.cs"]
    };

    caller_file_path_and_member_name_together => {
        r#"class Trace {
    public static void Show(
        [System.Runtime.CompilerServices.CallerMemberName] string member = "",
        [System.Runtime.CompilerServices.CallerFilePath] string path = "") {
        Console.WriteLine(member);
        Console.WriteLine(path.Length > 0);
    }
}
class App { public void Go() { Trace.Show(); } }
new App().Go();"#,
        ["Go", "True"]
    };

    caller_all_three_attributes_combined => {
        r#"class Trace {
    public static void Show(
        [System.Runtime.CompilerServices.CallerMemberName] string member = "",
        [System.Runtime.CompilerServices.CallerLineNumber] int line = 0,
        [System.Runtime.CompilerServices.CallerFilePath] string path = "") {
        Console.WriteLine(member);
        Console.WriteLine(line);
        Console.WriteLine(path.Length > 0);
    }
}
class App { public void Run() { Trace.Show(); } }
new App().Run();"#,
        ["Run", "8", "True"]
    };

    caller_member_name_on_static_property_getter => {
        r#"class Config {
    static int _port = 80;
    public static int Port {
        get {
            Log();
            return _port;
        }
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Console.WriteLine(Config.Port);"#,
        ["Port", "80"]
    };

    caller_member_name_on_static_property_setter => {
        r#"class Config {
    static int _port;
    public static int Port {
        set {
            Log();
            _port = value;
        }
        get => _port;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Config.Port = 443; Console.WriteLine(Config.Port);"#,
        ["Port", "443"]
    };

    caller_member_name_from_struct_method => {
        r#"struct Worker {
    public void DoWork() { Trace.Show(); }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Worker().DoWork();"#,
        ["DoWork"]
    };

    caller_member_name_from_interface_implementation => {
        r#"interface IRun { void Run(); }
class Job : IRun {
    public void Run() { Trace.Show(); }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
IRun job = new Job(); job.Run();"#,
        ["Run"]
    };

    caller_member_name_from_override_method => {
        r#"class Base { public virtual void Work() { } }
class Derived : Base {
    public override void Work() { Trace.Show(); }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Derived().Work();"#,
        ["Work"]
    };

    caller_line_number_from_loop_body => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
for (int i = 0; i < 1; i++) Trace.Show();"#,
        ["4"]
    };

    caller_line_number_from_switch_case => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
switch (1) {
    case 1: Trace.Show(); break;
}"#,
        ["5"]
    };

    caller_member_name_from_finally_block => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
try { int x = 1; } finally { Trace.Show(); }"#,
        ["<Main>$"]
    };

    caller_member_name_from_catch_clause => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
try { throw new System.Exception("x"); }
catch (System.Exception) { Trace.Show(); }"#,
        ["<Main>$"]
    };

    caller_member_name_on_private_helper => {
        r#"class Service {
    public int Compute() => Helper();
    int Helper([System.Runtime.CompilerServices.CallerMemberName] string member = "") {
        Console.WriteLine(member);
        return 1;
    }
}
Console.WriteLine(new Service().Compute());"#,
        ["Compute", "1"]
    };

    caller_member_name_on_extension_like_static => {
        r#"static class Ext {
    public static void Dump(this string s, [System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
"hi".Dump();"#,
        ["<Main>$"]
    };

    caller_line_number_on_delegate_invocation => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
System.Action act = () => Trace.Show();
act();"#,
        ["5"]
    };

    caller_member_name_multiple_optional_params_only_first_is_caller => {
        r#"class Trace {
    public static void Show(string prefix = "p", [System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(prefix + member);
}
class App { public void Run() { Trace.Show(); } }
new App().Run();"#,
        ["pRun"]
    };

    caller_member_name_on_generic_method => {
        r#"class Box {
    public T Read<T>([System.Runtime.CompilerServices.CallerMemberName] string member = "") {
        Console.WriteLine(member);
        return default(T);
    }
}
new Box().Read<int>();"#,
        ["Read"]
    };

    caller_file_path_contains_cs_suffix_when_explicit => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerFilePath] string path = "") => Console.WriteLine(path.EndsWith(".cs"));
}
Trace.Show("Program.cs");"#,
        ["True"]
    };

    caller_line_number_zero_when_explicit_zero => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = -1) => Console.WriteLine(line);
}
Trace.Show(0);"#,
        ["0"]
    };

    caller_member_name_empty_string_when_explicit_empty => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "x") => Console.WriteLine(member);
}
Trace.Show("");"#,
        [""]
    };

    caller_member_name_from_async_style_method_name => {
        r#"class Loader {
    public System.Threading.Tasks.Task<int> LoadAsync() {
        Trace.Show();
        return System.Threading.Tasks.Task.FromResult(1);
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
var t = new Loader().LoadAsync();
Console.WriteLine(t.Result);"#,
        ["LoadAsync", "1"]
    };

    caller_line_number_from_nested_block => {
        r#"class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => Console.WriteLine(line);
}
{
    Trace.Show();
}"#,
        ["5"]
    };

    caller_member_name_on_record_method => {
        r#"record Point(int X, int Y) {
    public void Report() {
        Trace.Show();
    }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
new Point(1, 2).Report();"#,
        ["Report"]
    };

    caller_all_three_respect_explicit_overrides => {
        r#"class Trace {
    public static void Show(
        [System.Runtime.CompilerServices.CallerMemberName] string member = "",
        [System.Runtime.CompilerServices.CallerLineNumber] int line = 0,
        [System.Runtime.CompilerServices.CallerFilePath] string path = "") {
        Console.WriteLine(member);
        Console.WriteLine(line);
        Console.WriteLine(path);
    }
}
Trace.Show("m", 42, "/a/b.cs");"#,
        ["m", "42", "/a/b.cs"]
    };

    caller_member_name_from_static_local_context => {
        r#"class MathUtil {
    public static int Square(int n) {
        Log();
        return n * n;
    }
    static void Log([System.Runtime.CompilerServices.CallerMemberName] string member = "") => Console.WriteLine(member);
}
Console.WriteLine(MathUtil.Square(4));"#,
        ["Square", "16"]
    };
}
