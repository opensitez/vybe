// vybe-test: csharp/csharp_linq_groupby_join/join_links_two_sequences_on_matching_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

var ids  = new[] { 1, 2, 3 };
var names = new[] { (Id:1, Name:"one"), (Id:2, Name:"two") };
var joined = ids.Join(names, id => id, n => n.Id, (id, n) => n.Name);
foreach (var s in joined) Console.WriteLine(s);
