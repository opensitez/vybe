// vybe-test: csharp/csharp_yield_iterators_core/yield_break_inside_nested_loop
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Grid(int rows,int cols){for(int r=0;r<rows;r++){for(int c=0;c<cols;c++){if(r==1&&c==1)yield break;yield return r*10+c;}}}
Console.WriteLine(string.Join(",",Grid(3,3)));
