// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_enum_like_factory
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

interface IKind<T> where T:IKind<T>{static abstract T North(); static abstract T South();}
struct Dir:IKind<Dir>{public string Name; public static Dir North()=>new Dir{Name="N"}; public static Dir South()=>new Dir{Name="S"};}
__P((Dir.North().Name).ToString());
__Check("N");
