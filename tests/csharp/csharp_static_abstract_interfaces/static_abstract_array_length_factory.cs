// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_array_length_factory
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IArray<T> where T:IArray<T>{static abstract int Length(T value);}
struct Arr:IArray<Arr>{public int[] Data; public static int Length(Arr value)=>value.Data.Length;}
__Check((Arr.Length(new Arr{Data=new[]{1,2,3}})).ToString(), "3");
