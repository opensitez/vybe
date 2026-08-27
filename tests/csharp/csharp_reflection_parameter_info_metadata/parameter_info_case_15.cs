// vybe-test: csharp/csharp_reflection_parameter_info_metadata/parameter_info_case_15

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

var method = typeof(Math).GetMethod("Max", new Type[] { typeof(int), typeof(int) });
var parameters = method.GetParameters();
__P(parameters.Length.ToString());
__P(parameters[0].ParameterType.Name);
__Check("2\nInt32");
