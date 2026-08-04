// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_override_calls_base_default_via_super_not_available_use_explicit
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

interface IA{string V()=>"a";} interface IB:IA{string W()=>V()+"b";} class Z:IB{public string V()=>"z";} __P((((IB)new Z()).W()).ToString());
__Check("ab");
