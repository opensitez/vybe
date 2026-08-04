// vybe-test: csharp/common_patterns/enum_with_values
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

enum Status { Active = 1, Inactive = 0, Pending = 2 }
__P(((int)Status.Active).ToString());
__P(((int)Status.Inactive).ToString());
__P(((int)Status.Pending).ToString());
__Check("1\n0\n2");
