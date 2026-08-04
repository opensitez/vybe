// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_chain_three_levels
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

interface I1{string S()=>"a";} interface I2:I1{string T()=>S()+"b";} class X:I2{} __P((new X().T()).ToString());
__Check("ab");
