// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_even_if_no_yield_reached
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

System.Collections.Generic.IEnumerable<int> Gen(bool ok){try{if(!ok)yield break;yield return 1;}finally{__P(("end").ToString());}}
foreach(var _ in Gen(false)){}
__Check("end");
