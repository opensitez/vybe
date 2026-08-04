// vybe-test: csharp/csharp_yield_iterators_core/yield_return_lazy_body_not_run_until_move_next
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

int calls=0; System.Collections.Generic.IEnumerable<int> Lazy(){calls++;yield return 1;}
var seq=Lazy(); __P((calls).ToString()); foreach(var _ in seq){} __P((calls).ToString());
__Check("0\n1");
