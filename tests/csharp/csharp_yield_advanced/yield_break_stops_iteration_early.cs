// vybe-test: csharp/csharp_yield_advanced/yield_break_stops_iteration_early
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

System.Collections.Generic.IEnumerable<int> Take(int[] a,int max){
    int count=0;
    foreach(var n in a){
        if(count>=max) yield break;
        yield return n;
        count++;
    }
}
Console.WriteLine(string.Join(",",Take(new[]{1,2,3,4,5},3)));
