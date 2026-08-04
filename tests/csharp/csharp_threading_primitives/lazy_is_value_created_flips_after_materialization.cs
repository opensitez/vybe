// vybe-test: csharp/csharp_threading_primitives/lazy_is_value_created_flips_after_materialization
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

var lazy = new System.Lazy<int>(() => 3);
__P((lazy.IsValueCreated).ToString());
__P((lazy.Value).ToString());
__P((lazy.IsValueCreated).ToString());
__Check("False\n3\nTrue");
