// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_passes_nested_to_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Store().Check()).ToString());
__Check("44");

class Store{public class Item{public int Id;} int Inspect(Item i)=>i.Id; public int Check()=>Inspect(new Item{Id=44});}

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
