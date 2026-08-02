// vybe-test: csharp/csharp_generics_constraints/generic_class_with_constraint_can_store_value
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder<T> where T : class { public T Value { get; set; } } var holder = new Holder<string> { Value = "abc" }; __Check((holder.Value).ToString(), "abc");
