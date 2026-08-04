// vybe-test: csharp/csharp_datetime_advanced/datetime_subtract_yields_timespan
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

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

var a=new System.DateTime(2024,1,10);
var b=new System.DateTime(2024,1,1);
var diff=a-b;
__P((diff.Days).ToString());
__Check("9");
