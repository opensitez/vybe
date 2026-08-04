// vybe-test: csharp/csharp_generics_advanced/default_of_generic_t_is_zero_for_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

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

T Zero<T>() => default(T);
__P((Zero<int>()).ToString());
__P((Zero<bool>()).ToString());
__Check("0\nFalse");
