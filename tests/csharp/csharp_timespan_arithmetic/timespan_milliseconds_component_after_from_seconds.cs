// vybe-test: csharp/csharp_timespan_arithmetic/timespan_milliseconds_component_after_from_seconds
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

var span=System.TimeSpan.FromSeconds(1.5); __P((span.Milliseconds).ToString());
__Check("500");
