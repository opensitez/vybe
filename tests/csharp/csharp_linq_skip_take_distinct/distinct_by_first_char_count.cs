// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_first_char_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

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

var r=new[]{"cat","car","dog","dot"}.DistinctBy(s=>s[0]);
__P((r.Count()).ToString());
__Check("2");
