// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_with_outer_generic_constraint
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Repo<int>().Read(new Repo<int>.Row{Data=77})).ToString());
__Check("77");

class Repo<T>{public class Row{public T Data;} public T Read(Row r)=>r.Data;}

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
