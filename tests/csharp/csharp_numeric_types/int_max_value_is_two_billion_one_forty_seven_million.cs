// vybe-test: csharp/csharp_numeric_types/int_max_value_is_two_billion_one_forty_seven_million
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

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

__P((int.MaxValue).ToString());
__Check("2147483647");
