// vybe-test: csharp/csharp_loops/for_loop_counts_down_with_decrement
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

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

string r=""; for(int i=3;i>=1;i--) r+=i; __P((r).ToString());
__Check("321");
