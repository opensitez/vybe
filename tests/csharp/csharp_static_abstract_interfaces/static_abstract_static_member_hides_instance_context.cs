// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_static_member_hides_instance_context
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IStaticOnly<T> where T:IStaticOnly<T>{static abstract int Count();}
struct Tally:IStaticOnly<Tally>{public int N; public static int Count()=>3;}
__Check((Tally.Count()).ToString(), "3");
