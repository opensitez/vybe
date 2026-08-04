// vybe-test: csharp/csharp_generics_constraints/generic_method_can_swap_two_values_by_tuple_assignment
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

(T, T) Swap<T>(T left, T right) { (left, right) = (right, left); return (left, right); } var result = Swap(1, 9); __P((result.Item1).ToString()); __P((result.Item2).ToString());
__Check("9\n1");
