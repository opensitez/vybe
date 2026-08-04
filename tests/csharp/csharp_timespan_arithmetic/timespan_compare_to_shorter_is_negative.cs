// vybe-test: csharp/csharp_timespan_arithmetic/timespan_compare_to_shorter_is_negative
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

var left=System.TimeSpan.FromMinutes(1); var right=System.TimeSpan.FromMinutes(2); __P((left.CompareTo(right)).ToString());
__Check("-1");
