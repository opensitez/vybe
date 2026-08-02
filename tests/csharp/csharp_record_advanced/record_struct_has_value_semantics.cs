// vybe-test: csharp/csharp_record_advanced/record_struct_has_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Vec(int X,int Y);
var a=new Vec(1,2); var b=a; // copy
b=b with{X=99};
__Check((a.X).ToString(), "1");
