// vybe-test: csharp/csharp_reflection_property_getter_setter_dynamic/property_info_dynamic_case_20

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

var prop = typeof(string).GetProperty("Length");
int len = (int)prop.GetValue("Hello_20");
__P(len.ToString());
__Check("8");
