// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_one_class_overrides_other_uses_default
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

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

interface IScale{int Scale(int n)=>n;} class Plain:IScale{} class Double:IScale{public int Scale(int n)=>n*2;} __P((new Plain().Scale(5)+new Double().Scale(5)).ToString());
__Check("15");
