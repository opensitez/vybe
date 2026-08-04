// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_array_length_factory
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

interface IArray<T> where T:IArray<T>{static abstract int Length(T value);}
struct Arr:IArray<Arr>{public int[] Data; public static int Length(Arr value)=>value.Data.Length;}
__P((Arr.Length(new Arr{Data=new[]{1,2,3}})).ToString());
__Check("3");
