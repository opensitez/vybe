// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_switch_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Gate{public enum Mode{On,Off} public string Label(Mode m){switch(m){case Mode.On:return "on"; default:return "off";}}} __P((new Gate().Label(Gate.Mode.On)).ToString());
__Check("on");
