// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_two_defaults_resolved_by_class_public_override
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

interface IA{void M()=>__P(("A").ToString());} interface IB{void M()=>__P(("B").ToString());} class C:IA,IB{public void M()=>__P(("C").ToString());} new C().M();
__Check("C");
