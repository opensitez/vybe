// vybe-test: csharp/csharp_readonly_members/readonly_struct_fields_all_readonly_by_definition
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Point{public readonly int X,Y; public Point(int x,int y){X=x;Y=y;}}
var p=new Point(1,2);
__P((p.X+p.Y).ToString());
__Check("3");
