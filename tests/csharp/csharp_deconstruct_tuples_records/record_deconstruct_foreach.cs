// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_foreach
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

record Point(int X,int Y); var pts=new[]{new Point(1,2),new Point(3,4)}; int sum=0; foreach(var (x,y) in pts) sum+=x+y; Console.WriteLine(sum);
