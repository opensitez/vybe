// vybe-test: csharp/csharp_threading_primitives/lazy_factory_runs_once_on_first_value_access
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

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

int calls = 0;
var lazy = new System.Lazy<int>(() => { calls++; return 7; });
__P((calls).ToString());
__P((lazy.Value).ToString());
__P((calls).ToString());
__Check("0\n7\n1");
