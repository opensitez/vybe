// vybe-test: csharp/csharp_linq_let_join/join_correlates_two_sequences_on_matching_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

var ids=new[]{1,2,3};
var labels=new[]{(Id:1,Text:"one"),(Id:2,Text:"two")};
var q=from id in ids
      join l in labels on id equals l.Id
      select l.Text;
foreach(var x in q) Console.WriteLine(x);
