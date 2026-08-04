// vybe-test: csharp/csharp_linq_aggregates/all_returns_false_when_one_element_fails_predicate
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

__P((new[]{2,4,5}.All(x => x%2==0)).ToString());
__Check("False");
