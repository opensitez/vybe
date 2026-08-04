// vybe-test: csharp/csharp_linq_skip_take_distinct/take_while_strings_prefix_length_le_two_count
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

var r=new[]{"a","bb","ccc"}.TakeWhile(s=>s.Length<=2);
__P((r.Count()).ToString());
__Check("2");
