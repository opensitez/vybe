// vybe-test: csharp/csharp_checked_overflow_throws/checked_add_on_int_max_value_throws_overflow_exception
// origin: languages/csharp/tests/csharp/test_csharp_checked_overflow_throws.rs

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

string outcome = "ok";
try {
    checked {
        int value = int.MaxValue;
        value += 1;
    }
} catch (System.OverflowException) {
    outcome = "overflow";
}
__P((outcome).ToString());
__Check("overflow");
