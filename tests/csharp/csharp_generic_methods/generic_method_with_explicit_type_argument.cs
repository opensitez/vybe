// vybe-test: csharp/csharp_generic_methods/generic_method_with_explicit_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

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

T Box<T>(T v)=>v;
__P((Box<int>(5)).ToString());
__Check("5");
