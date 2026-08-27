// vybe-test: csharp/csharp_generics_where/where_new_constraint_allows_parameterless_construction
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

T Build<T>() where T:new()=>new T();
__P((Build<Box>().V).ToString());
__Check("7");

class Box{public int V=7;}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
