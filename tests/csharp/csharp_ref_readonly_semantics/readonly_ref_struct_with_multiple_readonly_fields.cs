// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_with_multiple_readonly_fields
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly ref struct Rect{public readonly int W; public readonly int H; public Rect(int w,int h){W=w; H=h;} public int Area()=>W*H;} var r=new Rect(3,4); __Check((r.Area()).ToString(), "12");
