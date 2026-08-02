// vybe-test: csharp/csharp_operators/user_defined_implicit_conversion_coerces_to_target_type
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Inch {
    public double Value;
    public static implicit operator double(Inch i) => i.Value;
}
double length = new Inch { Value = 2.5 };
__Check((length).ToString(), "2.5");
