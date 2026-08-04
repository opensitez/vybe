// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_nested_interface_hierarchy
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IRoot<T> where T:IRoot<T>{static abstract T Root();}
interface IChild<T>:IRoot<T> where T:IChild<T>{static abstract T Child();}
struct Tree:IChild<Tree>{public string Tag; public static Tree Root()=>new Tree{Tag="R"}; public static Tree Child()=>new Tree{Tag="C"};}
__P((Tree.Child().Tag).ToString());
__Check("C");
