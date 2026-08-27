// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_via_skip_take_page_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var src=new[]{1,2,3,4,5,6}
;
int pageSize=2;
int pageCount=0;
for(int i=0;i<src.Length;i+=pageSize) pageCount++;
__P((pageCount).ToString());
__Check("3");

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
