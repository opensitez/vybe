// vybe-test: csharp/csharp_operator_overloading/unary_negation_operator_flips_sign_of_fields
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X;
public static Vec operator-(Vec v)=>new Vec{X=-v.X};}
var v=-new Vec{X=7};
__Check((v.X).ToString(), "-7");
