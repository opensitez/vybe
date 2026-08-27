// vybe-test: csharp/csharp_array_advanced/array_sort_then_binary_search_finds_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

using static __Harness;

int[] arr={5,3,1,4,2}
;
System.Array.Sort(arr);
int idx=System.Array.BinarySearch(arr,4);
__P((idx).ToString());
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
