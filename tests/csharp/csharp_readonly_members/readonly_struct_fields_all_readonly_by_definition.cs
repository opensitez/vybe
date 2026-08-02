// vybe-test: csharp/csharp_readonly_members/readonly_struct_fields_all_readonly_by_definition
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Point{public readonly int X,Y; public Point(int x,int y){X=x;Y=y;}}
var p=new Point(1,2);
__Check((p.X+p.Y).ToString(), "3");
