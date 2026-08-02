// vybe-test: csharp/csharp_struct_advanced/struct_default_keyword_produces_zero_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X,Y,Z;}
var v=default(Vec);
__Check((v.X==0&&v.Y==0&&v.Z==0).ToString(), "True");
