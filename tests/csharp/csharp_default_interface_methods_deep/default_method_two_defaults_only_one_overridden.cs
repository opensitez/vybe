// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_two_defaults_only_one_overridden
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

interface IA{int A()=>1;} interface IB{int B()=>2;} class Mix:IA,IB{public int A()=>10;} var m=new Mix(); __P((m.A()+m.B()).ToString());
__Check("12");
