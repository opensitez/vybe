// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_manual_window_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var src=new[]{10,20,30,40,50}
;
int size=2;
int windows=0;
for(int i=0;i+size<=src.Length;i+=size) windows++;
__P((windows).ToString());
__Check("2");

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
