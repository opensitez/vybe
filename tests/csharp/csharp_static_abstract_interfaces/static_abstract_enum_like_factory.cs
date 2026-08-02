// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_enum_like_factory
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IKind<T> where T:IKind<T>{static abstract T North(); static abstract T South();}
struct Dir:IKind<Dir>{public string Name; public static Dir North()=>new Dir{Name="N"}; public static Dir South()=>new Dir{Name="S"};}
__Check((Dir.North().Name).ToString(), "N");
