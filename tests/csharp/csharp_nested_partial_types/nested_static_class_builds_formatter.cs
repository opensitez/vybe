// vybe-test: csharp/csharp_nested_partial_types/nested_static_class_builds_formatter
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

__P((Report.Formatter.Line("count", 3)).ToString());
__Check("count:3");

class Report {
    public static class Formatter {
        public static string Line(string key, int value) { return key + ":" + value; }
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
