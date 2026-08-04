// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_with_struct_and_class_implementors
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

interface IShared<T> where T:IShared<T>{static abstract int Key();}
struct SA:IShared<SA>{public static int Key()=>1;} class CA:IShared<CA>{public static int Key()=>2;}
__P((SA.Key()+CA.Key()).ToString());
__Check("3");
