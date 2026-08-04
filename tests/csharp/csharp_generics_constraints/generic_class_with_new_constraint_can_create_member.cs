// vybe-test: csharp/csharp_generics_constraints/generic_class_with_new_constraint_can_create_member
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

class Factory<T> where T : new() { public T Build() { return new T(); } } class Item { public string Name = "built"; } __P((new Factory<Item>().Build().Name).ToString());
__Check("built");
