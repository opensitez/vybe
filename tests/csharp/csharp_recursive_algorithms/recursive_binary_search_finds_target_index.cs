// vybe-test: csharp/csharp_recursive_algorithms/recursive_binary_search_finds_target_index
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int BinSearch(int[] a,int lo,int hi,int t){
    if(lo>hi) return -1;
    int mid=(lo+hi)/2;
    return a[mid]==t?mid:a[mid]<t?BinSearch(a,mid+1,hi,t):BinSearch(a,lo,mid-1,t);
}
var arr=new[]{1,3,5,7,9,11};
__Check((BinSearch(arr,0,arr.Length-1,7)).ToString(), "3");
