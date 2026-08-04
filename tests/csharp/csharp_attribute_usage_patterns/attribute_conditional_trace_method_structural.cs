// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_trace_method_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; using System.Diagnostics; class Log{[Conditional("TRACE")] public static void Mark(){__P(("mark").ToString());} public static void Run(){Mark(); __P(("after").ToString());}} Log.Run();
__Check("after");
