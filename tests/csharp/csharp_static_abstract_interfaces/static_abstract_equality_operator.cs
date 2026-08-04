// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_equality_operator
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

interface IEquatableStatic<T> where T:IEquatableStatic<T>{static abstract bool operator==(T a,T b);}
struct Key:IEquatableStatic<Key>{public int Id; public static bool operator==(Key a,Key b)=>a.Id==b.Id; public static bool operator!=(Key a,Key b)=>!(a==b);}
__P((new Key{Id=1}==new Key{Id=1}).ToString());
__Check("True");
