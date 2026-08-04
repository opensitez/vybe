// vybe-test: csharp/csharp_delegate_types/predicate_t_tests_condition_on_value
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

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

System.Predicate<string> isLong = s => s.Length > 4;
__P((isLong("hello")).ToString());
__P((isLong("hi")).ToString());
__Check("True\nFalse");
