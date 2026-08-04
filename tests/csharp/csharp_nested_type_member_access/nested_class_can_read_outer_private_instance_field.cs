// vybe-test: csharp/csharp_nested_type_member_access/nested_class_can_read_outer_private_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

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

class Outer {
    int secret = 8;
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Read() { return parent.secret; }
    }
    public int ViaInner() { return new Inner(this).Read(); }
}
__P((new Outer().ViaInner()).ToString());
__Check("8");
