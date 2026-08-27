// vybe-test: csharp/csharp_patterns_extended_property_patterns/extended_property_case_3

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

var item = new { Info = new { Code = 3 } };
bool matches = item is { Info.Code: 3 };
__P(matches.ToString());
__Check("True");
