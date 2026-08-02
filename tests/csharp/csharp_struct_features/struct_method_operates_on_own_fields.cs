// vybe-test: csharp/csharp_struct_features/struct_method_operates_on_own_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vector { public double X,Y; public double Length() => System.Math.Sqrt(X*X+Y*Y); }
var v = new Vector { X=3, Y=4 };
__Check((v.Length()).ToString(), "5");
