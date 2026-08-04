// vybe-test: csharp/csharp_timespan_arithmetic/timespan_constructor_days_hours_minutes
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

var span=new System.TimeSpan(1,2,30); __P((span.Days).ToString()); __P((span.Hours).ToString()); __P((span.Minutes).ToString());
__Check("1\n2\n30");
