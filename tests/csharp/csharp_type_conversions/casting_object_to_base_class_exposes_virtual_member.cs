// vybe-test: csharp/csharp_type_conversions/casting_object_to_base_class_exposes_virtual_member
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public virtual string Name() { return "base"; } } class Child : Base { public override string Name() { return "child"; } } object item = new Child(); __Check((((Base)item).Name()).ToString(), "child");
