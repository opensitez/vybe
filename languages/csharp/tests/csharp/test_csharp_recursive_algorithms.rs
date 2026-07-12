//! Classic recursive algorithms as C# method implementations.
use super::helpers::run_csharp;

#[test]
fn recursive_factorial_returns_correct_result() {
    assert_eq!(
        run_csharp(
            r#"long Fact(int n)=>n<=1?1:n*Fact(n-1);
Console.WriteLine(Fact(10));"#
        ),
        &["3628800"]
    );
}

#[test]
fn recursive_fibonacci_returns_correct_nth_number() {
    assert_eq!(
        run_csharp(
            r#"int Fib(int n)=>n<=1?n:Fib(n-1)+Fib(n-2);
Console.WriteLine(Fib(8));"#
        ),
        &["21"]
    );
}

#[test]
fn recursive_gcd_computes_greatest_common_divisor() {
    assert_eq!(
        run_csharp(
            r#"int Gcd(int a,int b)=>b==0?a:Gcd(b,a%b);
Console.WriteLine(Gcd(48,18));"#
        ),
        &["6"]
    );
}

#[test]
fn recursive_binary_search_finds_target_index() {
    assert_eq!(
        run_csharp(
            r#"int BinSearch(int[] a,int lo,int hi,int t){
    if(lo>hi) return -1;
    int mid=(lo+hi)/2;
    return a[mid]==t?mid:a[mid]<t?BinSearch(a,mid+1,hi,t):BinSearch(a,lo,mid-1,t);
}
var arr=new[]{1,3,5,7,9,11};
Console.WriteLine(BinSearch(arr,0,arr.Length-1,7));"#
        ),
        &["3"]
    );
}

#[test]
fn recursive_quicksort_sorts_array_in_place() {
    assert_eq!(
        run_csharp(
            r#"void QSort(int[] a,int lo,int hi){
    if(lo>=hi) return;
    int p=a[hi],i=lo;
    for(int j=lo;j<hi;j++) if(a[j]<=p){int t=a[i];a[i]=a[j];a[j]=t;i++;}
    int tmp=a[i];a[i]=a[hi];a[hi]=tmp;
    QSort(a,lo,i-1); QSort(a,i+1,hi);
}
int[] arr={5,3,8,1,4};
QSort(arr,0,arr.Length-1);
Console.WriteLine(string.Join(",",arr));"#
        ),
        &["1,3,4,5,8"]
    );
}

#[test]
fn mutual_recursion_even_odd_check() {
    assert_eq!(
        run_csharp(
            r#"bool IsEven(int n){if(n==0)return true; return IsOdd(n-1);}
bool IsOdd(int n){if(n==0)return false; return IsEven(n-1);}
Console.WriteLine(IsEven(4)); Console.WriteLine(IsOdd(3));"#
        ),
        &["True", "True"]
    );
}
