// vybe-test: csharp/csharp_pattern_deconstruct/list_pattern_matches_exact_element_sequence
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

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

int[] data = { 1, 2, 3 };
if (data is [1, 2, 3]) __P(("exact").ToString());
else __P(("no").ToString());
__Check("exact");
