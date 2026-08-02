// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_record_style_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Orders{public class Line{public int Qty; public int Total()=>Qty*2;} public Line Make(int q)=>new Line{Qty=q};} __Check((new Orders().Make(4).Total()).ToString(), "8");
