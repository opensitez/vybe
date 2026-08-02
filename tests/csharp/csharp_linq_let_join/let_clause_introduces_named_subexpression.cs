// vybe-test: csharp/csharp_linq_let_join/let_clause_introduces_named_subexpression
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

var result =
    from s in new[]{"hello","hi","world"}
    let len=s.Length
    where len>3
    select s;
foreach(var x in result) Console.WriteLine(x);
