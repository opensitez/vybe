// vybe-test: csharp/csharp_enum_operations/enum_is_defined_returns_false_for_out_of_range_int
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

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

enum Level{Low=0,Mid=1,High=2}
__P((System.Enum.IsDefined(typeof(Level), 99)).ToString());
__Check("False");
