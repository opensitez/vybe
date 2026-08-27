// vybe-test: csharp/csharp_recursive_algorithms/recursive_quicksort_sorts_array_in_place
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

using static __Harness;

void QSort(int[] a,int lo,int hi){
    if(lo>=hi) return;
    int p=a[hi],i=lo;
    for(int j=lo;j<hi;j++) if(a[j]<=p){int t=a[i];a[i]=a[j];a[j]=t;i++;}
    int tmp=a[i];a[i]=a[hi];a[hi]=tmp;
    QSort(a,lo,i-1); QSort(a,i+1,hi);
}
int[] arr={5,3,8,1,4}
;
QSort(arr,0,arr.Length-1);
__P((string.Join(",",arr)).ToString());
__Check("1,3,4,5,8");

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
