// vybe-test: csharp/csharp_linq_aggregates/contains_returns_true_for_present_value_in_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

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

__P((new[]{1,2,3}.Contains(2)).ToString());
__Check("True");
