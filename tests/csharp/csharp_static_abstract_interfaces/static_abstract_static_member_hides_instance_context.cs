// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_static_member_hides_instance_context
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

interface IStaticOnly<T> where T:IStaticOnly<T>{static abstract int Count();}
struct Tally:IStaticOnly<Tally>{public int N; public static int Count()=>3;}
__P((Tally.Count()).ToString());
__Check("3");
