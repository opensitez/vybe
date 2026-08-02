// vybe-test: csharp/csharp_generics_constraints/generic_class_with_base_constraint_can_call_virtual_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public virtual string Read() { return "base"; } } class Child : Base { public override string Read() { return "child"; } } class Reader<T> where T : Base { public string Run(T value) { return value.Read(); } } __Check((new Reader<Child>().Run(new Child())).ToString(), "child");
