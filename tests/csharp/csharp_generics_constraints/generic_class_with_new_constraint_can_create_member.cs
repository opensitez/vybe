// vybe-test: csharp/csharp_generics_constraints/generic_class_with_new_constraint_can_create_member
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Factory<T> where T : new() { public T Build() { return new T(); } } class Item { public string Name = "built"; } __Check((new Factory<Item>().Build().Name).ToString(), "built");
