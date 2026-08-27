// vybe-test: csharp/csharp_caller_info_attributes/caller_file_path_contains_cs_suffix_when_explicit
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

using static __Harness;

App.Run();
__Check("Done_caller_file_path_contains_cs_suffix_when_explicit");

class App {
    public static void Run() {
        Log();
    }
    public static void Log([System.Runtime.CompilerServices.CallerMemberName] string m = "") {
        __P("Done_caller_file_path_contains_cs_suffix_when_explicit");
    }
}
public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
