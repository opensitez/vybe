// vybe-test: csharp/csharp_datetime_dateonly_arithmetic/dateonly_operations_case_16

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var d = new DateOnly(2026, 5, 16);
var next = d.AddDays(10);
__P(d.Day.ToString());
__P(next.Day.ToString());
__Check("16\n26");
