// vybe-test: csharp/csharp_struct_advanced/struct_passed_by_in_not_copied_but_read_only
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

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

struct Vec{public int X,Y;}
int Sum(in Vec v)=>v.X+v.Y;
var v=new Vec{X=3,Y=4};
__P((Sum(in v)).ToString());
__Check("7");
