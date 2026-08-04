// vybe-test: csharp/csharp_enum_flags_operations/enum_parse_reads_name_into_typed_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

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

enum Color { Red, Green, Blue }
var value = System.Enum.Parse(typeof(Color), "Green");
__P((value).ToString());
__Check("Green");
