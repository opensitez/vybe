// vybe-test: csharp/csharp_recursive_algorithms/recursive_binary_search_finds_target_index
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int BinSearch(int[] a,int lo,int hi,int t){
    if(lo>hi) return -1;
    int mid=(lo+hi)/2;
    return a[mid]==t?mid:a[mid]<t?BinSearch(a,mid+1,hi,t):BinSearch(a,lo,mid-1,t);
}
var arr=new[]{1,3,5,7,9,11};
__P((BinSearch(arr,0,arr.Length-1,7)).ToString());
__Check("3");
