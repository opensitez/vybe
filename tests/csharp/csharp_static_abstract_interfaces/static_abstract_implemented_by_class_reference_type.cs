// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_implemented_by_class_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBuild<T> where T:IBuild<T>{static abstract T New();}
class Node:IBuild<Node>{public int Id=7; public static Node New()=>new Node();}
__Check((Node.New().Id).ToString(), "7");
