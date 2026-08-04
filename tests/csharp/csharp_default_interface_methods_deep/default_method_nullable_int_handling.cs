// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_nullable_int_handling
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

interface INull{int? N{get;} int OrZero()=>N??0;} class Maybe:INull{public int? N{get;set;}=null;} __P((new Maybe().OrZero()).ToString());
__Check("0");
