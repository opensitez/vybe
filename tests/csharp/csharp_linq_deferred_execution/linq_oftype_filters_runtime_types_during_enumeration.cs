// vybe-test: csharp/csharp_linq_deferred_execution/linq_oftype_filters_runtime_types_during_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

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

using System.Linq;
object[] items = { 1, "a", 2, "b", 3 };
foreach (var text in items.OfType<string>()) __P((text).ToString());
__Check("a\nb");
