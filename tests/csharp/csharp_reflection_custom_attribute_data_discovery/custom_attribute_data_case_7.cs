// vybe-test: csharp/csharp_reflection_custom_attribute_data_discovery/custom_attribute_data_case_7

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

Type t = typeof(string);
var attrs = System.Reflection.CustomAttributeData.GetCustomAttributes(t);
__P((attrs != null).ToString());
__Check("True");
