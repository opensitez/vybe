// vybe-test: csharp/csharp_timespan_arithmetic/timespan_to_string_positive_hms
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

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

__P((System.TimeSpan.FromHours(1).Add(System.TimeSpan.FromMinutes(2)).Add(System.TimeSpan.FromSeconds(3)).ToString()).ToString());
__Check("01:02:03");
