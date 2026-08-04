// vybe-test: csharp/csharp_number_bases/convert_to_string_with_base_2_formats_binary
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

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

__P((System.Convert.ToString(10,2)).ToString());
__Check("1010");
