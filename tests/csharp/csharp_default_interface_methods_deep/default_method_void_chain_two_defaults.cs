// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_void_chain_two_defaults
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

interface IA{void A(){__P(("a").ToString());}} interface IB{void B(){__P(("b").ToString());}} class Both:IA,IB{} var b=new Both(); b.A(); b.B();
__Check("a\nb");
