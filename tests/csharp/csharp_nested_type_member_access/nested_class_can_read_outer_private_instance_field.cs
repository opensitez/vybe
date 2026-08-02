// vybe-test: csharp/csharp_nested_type_member_access/nested_class_can_read_outer_private_instance_field
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((new Outer().ViaInner()).ToString(), "8");
