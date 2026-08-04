// vybe-test: csharp/csharp_string_join_split/split_on_string_delimiter
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

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

var parts="one::two::three".Split("::");
__P((parts.Length).ToString()); __P((parts[2]).ToString());
__Check("3\nthree");
