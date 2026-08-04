// vybe-test: csharp/csharp_linq_materialization/linq_average_computes_mean_of_numeric_sequence
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
__P((new[] { 2, 4, 6 }.Average()).ToString());
__Check("4");
