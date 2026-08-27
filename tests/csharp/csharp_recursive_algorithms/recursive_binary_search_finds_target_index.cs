// vybe-test: csharp/csharp_recursive_algorithms/recursive_binary_search_finds_target_index
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

using static __Harness;

int BinSearch(int[] a,int lo,int hi,int t){
    if(lo>hi) return -1;
    int mid=(lo+hi)/2;
    return a[mid]==t?mid:a[mid]<t?BinSearch(a,mid+1,hi,t):BinSearch(a,lo,mid-1,t);
}
var arr=new[]{1,3,5,7,9,11}
;
__P((BinSearch(arr,0,arr.Length-1,7)).ToString());
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
