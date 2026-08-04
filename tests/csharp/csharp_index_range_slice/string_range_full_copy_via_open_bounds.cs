// vybe-test: csharp/csharp_index_range_slice/string_range_full_copy_via_open_bounds
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

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

string text="abc"; __P((text[..]).ToString());
__Check("abc");
