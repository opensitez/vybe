// vybe-test: csharp/csharp_reflection_activation/is_assignable_from_reports_true_for_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { } class Child : Base { } __Check((typeof(Base).IsAssignableFrom(typeof(Child))).ToString(), "True");
