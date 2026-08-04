// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_explicit_interface_implementation_style
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

interface IMaker<T> where T:IMaker<T>{static abstract T Make();}
struct Box:IMaker<Box>{public int Size; static Box IMaker<Box>.Make()=>new Box{Size=9}; public static Box Make()=>new Box{Size=1};}
__P((Box.Make().Size).ToString());
__Check("1");
