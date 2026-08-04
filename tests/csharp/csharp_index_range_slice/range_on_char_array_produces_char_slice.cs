// vybe-test: csharp/csharp_index_range_slice/range_on_char_array_produces_char_slice
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

char[] letters={'a','b','c','d'}; var slice=letters[1..3]; __P((slice.Length).ToString()); __P((slice[0]).ToString());
__Check("2\n98");
