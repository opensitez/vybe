// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_property_returns_field_by_readonly_ref
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point{public int X; public ref readonly int Rx=>ref X;} var p=new Point(); p.X=11; __Check((p.Rx).ToString(), "11");
