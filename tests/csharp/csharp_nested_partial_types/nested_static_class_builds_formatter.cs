// vybe-test: csharp/csharp_nested_partial_types/nested_static_class_builds_formatter
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Report {
    public static class Formatter {
        public static string Line(string key, int value) { return key + ":" + value; }
    }
}
__P((Report.Formatter.Line("count", 3)).ToString());
__Check("count:3");
