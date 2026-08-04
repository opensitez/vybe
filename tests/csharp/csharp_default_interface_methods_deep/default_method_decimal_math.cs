// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_decimal_math
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

interface IMoney{decimal Add(decimal a,decimal b)=>a+b;} class Wallet:IMoney{} __P((new Wallet().Add(1.5m,2.5m)).ToString());
__Check("4.0");
