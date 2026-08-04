// vybe-test: csharp/csharp_linq_materialization/linq_first_throws_on_empty_sequence_when_not_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

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
try {
    __P((new int[0].First()).ToString());
} catch (System.InvalidOperationException) {
    __P(("empty").ToString());
}
__Check("empty");
