// vybe-test: csharp/csharp_operator_overloading/implicit_conversion_from_int_to_custom_type
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Meters{public double Value;
public static implicit operator Meters(double d)=>new Meters{Value=d};}
Meters m=3.5;
__Check((m.Value).ToString(), "3.5");
