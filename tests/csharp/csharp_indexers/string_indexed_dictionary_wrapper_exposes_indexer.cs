// vybe-test: csharp/csharp_indexers/string_indexed_dictionary_wrapper_exposes_indexer
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

using static __Harness;

var c=new Config();
c["env"]="prod";
__P((c["env"]).ToString());
__Check("prod");

class Config{
    System.Collections.Generic.Dictionary<string,string> _d=new();
    public string this[string k]{get=>_d[k]; set=>_d[k]=value;}
}

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
