// vybe-test: csharp/csharp_operator_overloading/explicit_conversion_to_primitive_requires_cast
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Percent{public double Value;
public static explicit operator double(Percent p)=>p.Value/100.0;}
var p=new Percent{Value=50};
__Check(((double)p).ToString(), "0.5");
