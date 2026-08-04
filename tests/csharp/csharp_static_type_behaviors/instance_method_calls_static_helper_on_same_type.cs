// vybe-test: csharp/csharp_static_type_behaviors/instance_method_calls_static_helper_on_same_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Converter {
    public static int Double(int value) { return value * 2; }
    public int Convert(int value) { return Double(value) + 1; }
}
var converter = new Converter();
__P((converter.Convert(5)).ToString());
__Check("11");
