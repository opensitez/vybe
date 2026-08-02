// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_from_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Map{public struct Point{public int X; public int Y;} public Point Origin()=>new Point{X=0,Y=0};} var p=new Map().Origin(); __Check((p.X).ToString(), "0");
