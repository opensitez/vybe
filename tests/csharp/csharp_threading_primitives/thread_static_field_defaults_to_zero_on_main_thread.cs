// vybe-test: csharp/csharp_threading_primitives/thread_static_field_defaults_to_zero_on_main_thread
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

class Counter {
    [System.ThreadStatic]
    public static int Value;
}
__P((Counter.Value).ToString());
__Check("0");
