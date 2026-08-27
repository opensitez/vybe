// vybe-test: csharp/csharp_reflection_field_info_dynamic_access/field_info_case_17

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

var field = typeof(int).GetField("MaxValue");
int max = (int)field.GetValue(null);
__P((max == int.MaxValue).ToString());
__Check("True");
