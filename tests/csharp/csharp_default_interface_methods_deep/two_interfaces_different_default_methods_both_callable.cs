// vybe-test: csharp/csharp_default_interface_methods_deep/two_interfaces_different_default_methods_both_callable
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

interface IA{int A()=>1;} interface IB{int B()=>2;} class Both:IA,IB{} var x=new Both(); __P((x.A()+x.B()).ToString());
__Check("3");
