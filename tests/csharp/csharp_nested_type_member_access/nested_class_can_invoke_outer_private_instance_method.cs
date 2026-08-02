// vybe-test: csharp/csharp_nested_type_member_access/nested_class_can_invoke_outer_private_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    int Twice(int n) { return n * 2; }
    class Inner {
        Outer parent;
        public Inner(Outer parent) { this.parent = parent; }
        public int Run(int n) { return parent.Twice(n); }
    }
    public int ViaInner(int n) { return new Inner(this).Run(n); }
}
__Check((new Outer().ViaInner(5)).ToString(), "10");
