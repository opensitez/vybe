// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_type_params_with_self
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IPair<TSelf,TVal> where TSelf:IPair<TSelf,TVal>{static abstract TSelf Of(TVal v);}
struct Holder:IPair<Holder,int>{public int Data; public static Holder Of(int v)=>new Holder{Data=v};}
__Check((Holder.Of(6).Data).ToString(), "6");
