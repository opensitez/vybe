// vybe-test: csharp/csharp_string_parsing/enum_try_parse_succeeds_on_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

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

enum Color{Red,Green,Blue}
__P((System.Enum.TryParse<Color>("Green",out var c)).ToString());
__P((c).ToString());
__Check("True\nGreen");
