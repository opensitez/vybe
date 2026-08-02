// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_implemented_by_struct_value_type
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IVal<T> where T:IVal<T>{static abstract T Make(int n);}
struct Point:IVal<Point>{public int X; public static Point Make(int n)=>new Point{X=n};}
__Check((Point.Make(11).X).ToString(), "11");
