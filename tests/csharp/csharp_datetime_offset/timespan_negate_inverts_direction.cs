// vybe-test: csharp/csharp_datetime_offset/timespan_negate_inverts_direction
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

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

var ts=System.TimeSpan.FromHours(3);
__P(((-ts).Hours).ToString());
__Check("-3");
