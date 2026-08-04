// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_combined_if_and_conditional_print
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

#define VYBETEST_PRE
using System; using System.Diagnostics; class App{[Conditional("DEBUG")] static void Log(){} static void Main(){#if VYBETEST_PRE __P(("pre").ToString()); #endif Log(); __P(("post").ToString());}} App.Main();
__Check("pre\npost");
