// vybe-test: csharp/csharp_generics_constraints/generic_class_with_base_constraint_can_call_virtual_method
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

class Base { public virtual string Read() { return "base"; } } class Child : Base { public override string Read() { return "child"; } } class Reader<T> where T : Base { public string Run(T value) { return value.Read(); } } __P((new Reader<Child>().Run(new Child())).ToString());
__Check("child");
