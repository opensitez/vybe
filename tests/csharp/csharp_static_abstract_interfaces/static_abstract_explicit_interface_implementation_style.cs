// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_explicit_interface_implementation_style
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IMaker<T> where T:IMaker<T>{static abstract T Make();}
struct Box:IMaker<Box>{public int Size; static Box IMaker<Box>.Make()=>new Box{Size=9}; public static Box Make()=>new Box{Size=1};}
__Check((Box.Make().Size).ToString(), "1");
