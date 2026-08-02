// vybe-test: csharp/csharp_ienumerable_custom/custom_enumerable_iterated_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

class UpTo:System.Collections.Generic.IEnumerable<int>{
    int _max;
    public UpTo(int max){_max=max;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=1;i<=_max;i++) yield return i;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
int sum=0; foreach(var n in new UpTo(5)) sum+=n;
Console.WriteLine(sum);
