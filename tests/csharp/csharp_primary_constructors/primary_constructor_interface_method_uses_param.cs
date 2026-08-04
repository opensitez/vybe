// vybe-test: csharp/csharp_primary_constructors/primary_constructor_interface_method_uses_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

interface IVal { int Get(); }
class Impl(int n) : IVal { public int Get() => n; }
IVal v = new Impl(12);
__P((v.Get()).ToString());
__Check("12");
