// vybe-test: csharp/csharp_linq_numeric/long_count_works_on_large_range
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

long c=Enumerable.Range(0,1000).LongCount();
__P((c).ToString());
__Check("1000");
