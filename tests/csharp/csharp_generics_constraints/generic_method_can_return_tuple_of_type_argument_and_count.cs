// vybe-test: csharp/csharp_generics_constraints/generic_method_can_return_tuple_of_type_argument_and_count
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

(T, int) Pair<T>(T value) { return (value, 1); } var result = Pair("x"); __P((result.Item1 + result.Item2).ToString());
__Check("x1");
