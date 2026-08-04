// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_three_interfaces_class_unified_override
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

interface IA{void P()=>__P(("A").ToString());} interface IB{void P()=>__P(("B").ToString());} interface IC{void P()=>__P(("C").ToString());} class U:IA,IB,IC{public void P()=>__P(("U").ToString());} new U().P();
__Check("U");
