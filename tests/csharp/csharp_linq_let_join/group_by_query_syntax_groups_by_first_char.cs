// vybe-test: csharp/csharp_linq_let_join/group_by_query_syntax_groups_by_first_char
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

var words=new[]{"apple","ant","banana"};
var groups=from w in words group w by w[0];
int count=0;
foreach(var g in groups) count++;
Console.WriteLine(count);
