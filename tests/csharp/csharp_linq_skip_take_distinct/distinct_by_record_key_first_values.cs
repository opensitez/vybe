// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_record_key_first_values
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

var r=new[]{(K:1,V:"a"),(K:1,V:"b"),(K:2,V:"c")}.DistinctBy(t=>t.K);
__P((r.First().V).ToString()); __P((r.Last().V).ToString());
__Check("a\nc");
