// vybe-test: csharp/csharp_generics_constraints/generic_static_field_is_independent_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

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

class Counter<T> { public static int Value; } Counter<int>.Value = 2; Counter<string>.Value = 5; __P((Counter<int>.Value).ToString()); __P((Counter<string>.Value).ToString());
__Check("2\n5");
