// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_runs_once_per_enumeration
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

int fin=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin++;__P((fin).ToString());}}
foreach(var _ in Gen()){} foreach(var _ in Gen()){}
__Check("1\n2");
