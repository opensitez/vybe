// vybe-test: csharp/csharp_operator_overloading/multiply_operator_scales_vector_by_scalar
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X;
public static Vec operator*(Vec v,int s)=>new Vec{X=v.X*s};}
__Check(((new Vec{X=3}*4).X).ToString(), "12");
