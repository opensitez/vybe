// vybe-test: csharp/csharp_indexers/interface_defines_indexer_contract_implemented_by_class
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

using static __Harness;

IMap m=new Map();
__P((m[2]).ToString());
__Check("two");

interface IMap{string this[int k]{get;}}

class Map:IMap{string[] data={"zero","one","two"};public string this[int k]=>data[k];}

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
