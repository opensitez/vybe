// vybe-test: csharp/csharp_linq_deferred_execution/linq_second_enumeration_reexecutes_select_projection
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
int projections = 0;
var query = new[] { 5 }.Select(x => { projections++; return x + 1; });
__P((query.First()).ToString());
__P((projections).ToString());
__P((query.First()).ToString());
__P((projections).ToString());
__Check("6\n1\n6\n2");
