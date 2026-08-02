// vybe-test: csharp/csharp_ienumerable_custom/linq_works_on_custom_ienumerable
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

class Odds:System.Collections.Generic.IEnumerable<int>{
    int _count;
    public Odds(int count){_count=count;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=0;i<_count;i++) yield return 2*i+1;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
Console.WriteLine(new Odds(4).Sum());
