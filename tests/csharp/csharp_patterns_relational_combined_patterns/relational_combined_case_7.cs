// vybe-test: csharp/csharp_patterns_relational_combined_patterns/relational_combined_case_7

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

int val = 7;
bool inRange = val is > 0 and <= 20;
__P(inRange.ToString());
__Check("True");
