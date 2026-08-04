// vybe-test: csharp/csharp_string_methods/split_with_remove_empty_entries_drops_consecutive_delimiters
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

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

var p = "a,,b".Split(new[]{','}, System.StringSplitOptions.RemoveEmptyEntries);
__P((p.Length).ToString());
__Check("2");
