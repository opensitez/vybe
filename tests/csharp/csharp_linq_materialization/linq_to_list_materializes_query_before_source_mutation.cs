// vybe-test: csharp/csharp_linq_materialization/linq_to_list_materializes_query_before_source_mutation
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

using static __Harness;
using System.Collections.Generic;
using System.Linq;

var source = new List<int> { 1, 2 }
;
var snapshot = source.Select(x => x).ToList();
source.Add(3);
__P((snapshot.Count).ToString());
__P((source.Count).ToString());
__Check("2\n3");

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
