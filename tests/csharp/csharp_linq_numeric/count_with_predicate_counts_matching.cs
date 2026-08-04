// vybe-test: csharp/csharp_linq_numeric/count_with_predicate_counts_matching
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

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

__P((new[]{1,2,3,4,5,6}.Count(n=>n%2==0)).ToString());
__Check("3");
