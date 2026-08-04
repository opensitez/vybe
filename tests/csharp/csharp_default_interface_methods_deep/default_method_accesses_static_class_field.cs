// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_accesses_static_class_field
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

interface IStatic{int Read()=>Holder.N;} static class Holder{public static int N=8;} class R:IStatic{} __P((new R().Read()).ToString());
__Check("8");
