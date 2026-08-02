// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_used_in_generic_helper
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IShow<T> where T:IShow<T>{static abstract string Show(T v);}
struct Lab:IShow<Lab>{public int N; public static string Show(Lab v)=>v.N.ToString();}
string Render<T>(T v) where T:IShow<T>=>T.Show(v); __Check((Render(new Lab{N=9})).ToString(), "9");
