// vybe-test: csharp/csharp_generics_where/where_struct_constraint_accepts_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

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

T Default<T>() where T:struct=>default;
__P((Default<int>()).ToString());
__P((Default<bool>()).ToString());
__Check("0\nFalse");
