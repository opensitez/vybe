// vybe-test: csharp/csharp_string_methods/substring_extracts_region_by_start_and_length
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

__P(("hello world".Substring(6, 5)).ToString());
__Check("world");
