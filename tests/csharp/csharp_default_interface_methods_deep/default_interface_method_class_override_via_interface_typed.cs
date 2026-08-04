// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_method_class_override_via_interface_typed
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

interface IFormat{string Show(int n)=>n.ToString();} class Custom:IFormat{public string Show(int n)=>"x"+n;} IFormat f=new Custom(); __P((f.Show(3)).ToString());
__Check("x3");
