// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_static_method
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

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

class Seq{public static System.Collections.Generic.IEnumerable<int> Twice(int n){yield return n;yield return n*2;}}
__P((string.Join(",",Seq.Twice(5))).ToString());
__Check("5,10");
