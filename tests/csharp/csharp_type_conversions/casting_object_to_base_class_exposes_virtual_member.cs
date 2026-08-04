// vybe-test: csharp/csharp_type_conversions/casting_object_to_base_class_exposes_virtual_member
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

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

class Base { public virtual string Name() { return "base"; } } class Child : Base { public override string Name() { return "child"; } } object item = new Child(); __P((((Base)item).Name()).ToString());
__Check("child");
