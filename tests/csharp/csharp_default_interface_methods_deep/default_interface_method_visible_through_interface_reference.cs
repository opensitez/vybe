// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_method_visible_through_interface_reference
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

interface ICalc{int Double(int n)=>n*2;} class Worker:ICalc{} ICalc w=new Worker(); __P((w.Double(4)).ToString());
__Check("8");
