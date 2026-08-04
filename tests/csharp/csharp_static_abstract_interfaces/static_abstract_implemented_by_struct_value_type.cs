// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_implemented_by_struct_value_type
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface IVal<T> where T:IVal<T>{static abstract T Make(int n);}
struct Point:IVal<Point>{public int X; public static Point Make(int n)=>new Point{X=n};}
__P((Point.Make(11).X).ToString());
__Check("11");
