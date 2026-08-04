// vybe-test: csharp/csharp_generics_constraints/generic_method_with_base_class_constraint_accesses_base_member
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

class Base { public string Name = "base"; } class Child : Base { } string Read<T>(T value) where T : Base { return value.Name; } __P((Read(new Child())).ToString());
__Check("base");
