// vybe-test: csharp/csharp_patterns_positional_property_subpatterns/property_subpattern_case_15

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

string text = "Alpha_15";
bool isLong = text is { Length: >= 5 };
__P(isLong.ToString());
__Check("True");
