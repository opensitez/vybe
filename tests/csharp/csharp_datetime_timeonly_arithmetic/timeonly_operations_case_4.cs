// vybe-test: csharp/csharp_datetime_timeonly_arithmetic/timeonly_operations_case_4

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

var t = new TimeOnly(4, 30, 0);
var t2 = t.AddHours(2);
__P(t.Hour.ToString());
__P(t2.Hour.ToString());
__Check("4\n6");
