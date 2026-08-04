// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_two_defaults_resolved_by_explicit_interface_impl
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

interface IA{void M()=>__P(("A").ToString());} interface IB{void M()=>__P(("B").ToString());} class C:IA,IB{void IA.M()=>__P(("IA").ToString()); void IB.M()=>__P(("IB").ToString());} ((IA)new C()).M(); ((IB)new C()).M();
__Check("IA\nIB");
