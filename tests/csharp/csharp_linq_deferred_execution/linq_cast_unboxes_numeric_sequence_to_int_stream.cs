// vybe-test: csharp/csharp_linq_deferred_execution/linq_cast_unboxes_numeric_sequence_to_int_stream
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
object[] boxed = { 1, 2, 3 };
foreach (var value in boxed.Cast<int>()) __P((value + 1).ToString());
__Check("2\n3\n4");
