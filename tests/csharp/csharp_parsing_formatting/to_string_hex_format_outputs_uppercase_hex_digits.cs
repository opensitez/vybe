// vybe-test: csharp/csharp_parsing_formatting/to_string_hex_format_outputs_uppercase_hex_digits
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

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

__P((255.ToString("X")).ToString());
__Check("FF");
