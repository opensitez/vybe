// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_recursive_default_calls_itself
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

interface IRec{int Fact(int n)=>n<=1?1:n*Fact(n-1);} class Math:IRec{} __P((new Math().Fact(5)).ToString());
__Check("120");
