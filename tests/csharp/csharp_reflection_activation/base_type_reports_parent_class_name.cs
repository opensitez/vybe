// vybe-test: csharp/csharp_reflection_activation/base_type_reports_parent_class_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { } class Child : Base { } __Check((typeof(Child).BaseType.Name).ToString(), "Base");
