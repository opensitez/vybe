// vybe-test: csharp/csharp_interface_default_impl/two_classes_same_interface_one_overrides_one_uses_default
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

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

interface IFormat{string Format(int n)=>$"[{n}]";}
class A:IFormat{}
class B:IFormat{public string Format(int n)=>n.ToString();}
IFormat a=new A(); IFormat b=new B();
__P((a.Format(5)).ToString());
__P((b.Format(5)).ToString());
__Check("[5]\n5");
