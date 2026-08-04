// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_diamond_class_picks_single_public_impl
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

interface IA{void Print()=>__P(("A").ToString());} interface IB{void Print()=>__P(("B").ToString());} class Merge:IA,IB{public void Print()=>__P(("M").ToString());} new Merge().Print();
__Check("M");
