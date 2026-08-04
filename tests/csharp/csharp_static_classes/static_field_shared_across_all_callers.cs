// vybe-test: csharp/csharp_static_classes/static_field_shared_across_all_callers
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

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

static class Counter { public static int Count = 0; }
Counter.Count++;
Counter.Count++;
__P((Counter.Count).ToString());
__Check("2");
