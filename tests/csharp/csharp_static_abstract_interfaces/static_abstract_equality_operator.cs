// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_equality_operator
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IEquatableStatic<T> where T:IEquatableStatic<T>{static abstract bool operator==(T a,T b);}
struct Key:IEquatableStatic<Key>{public int Id; public static bool operator==(Key a,Key b)=>a.Id==b.Id; public static bool operator!=(Key a,Key b)=>!(a==b);}
__Check((new Key{Id=1}==new Key{Id=1}).ToString(), "True");
