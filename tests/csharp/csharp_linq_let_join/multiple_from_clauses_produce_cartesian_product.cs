// vybe-test: csharp/csharp_linq_let_join/multiple_from_clauses_produce_cartesian_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

var result=from a in new[]{1,2} from b in new[]{10,20} select a*b;
int sum=0; foreach(var x in result) sum+=x;
Console.WriteLine(sum);
