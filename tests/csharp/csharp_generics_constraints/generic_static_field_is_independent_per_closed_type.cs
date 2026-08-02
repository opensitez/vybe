// vybe-test: csharp/csharp_generics_constraints/generic_static_field_is_independent_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter<T> { public static int Value; } Counter<int>.Value = 2; Counter<string>.Value = 5; __Check((Counter<int>.Value).ToString(), "2"); __Check((Counter<string>.Value).ToString(), "5");
