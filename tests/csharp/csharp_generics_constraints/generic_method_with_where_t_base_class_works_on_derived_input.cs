// vybe-test: csharp/csharp_generics_constraints/generic_method_with_where_t_base_class_works_on_derived_input
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person { public string Name = "Ada"; } class Admin : Person { } string Read<T>(T person) where T : Person { return person.Name; } __Check((Read(new Admin())).ToString(), "Ada");
