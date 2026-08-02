// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_nested_interface_hierarchy
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRoot<T> where T:IRoot<T>{static abstract T Root();}
interface IChild<T>:IRoot<T> where T:IChild<T>{static abstract T Child();}
struct Tree:IChild<Tree>{public string Tag; public static Tree Root()=>new Tree{Tag="R"}; public static Tree Child()=>new Tree{Tag="C"};}
__Check((Tree.Child().Tag).ToString(), "C");
