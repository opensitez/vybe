// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_deconstruct_foreach_list
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var list=new System.Collections.Generic.List<Pair>{new Pair{A=1,B=2},new Pair{A=3,B=4}}; int sum=0; foreach(var (a,b) in list) sum+=a+b; Console.WriteLine(sum);
