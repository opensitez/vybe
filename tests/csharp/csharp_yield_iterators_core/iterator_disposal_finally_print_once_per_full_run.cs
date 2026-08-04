// vybe-test: csharp/csharp_yield_iterators_core/iterator_disposal_finally_print_once_per_full_run
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

int hits=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{hits++;__P((hits).ToString());}}
foreach(var _ in Gen()){} __P((hits).ToString());
__Check("1\n1");
