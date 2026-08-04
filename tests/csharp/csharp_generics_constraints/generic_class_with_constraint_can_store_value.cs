// vybe-test: csharp/csharp_generics_constraints/generic_class_with_constraint_can_store_value
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

class Holder<T> where T : class { public T Value { get; set; } } var holder = new Holder<string> { Value = "abc" }; __P((holder.Value).ToString());
__Check("abc");
