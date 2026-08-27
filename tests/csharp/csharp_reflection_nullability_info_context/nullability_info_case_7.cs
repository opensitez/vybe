// vybe-test: csharp/csharp_reflection_nullability_info_context/nullability_info_case_7

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

var ctx = new System.Reflection.NullabilityInfoContext();
__P((ctx != null).ToString());
__Check("True");
