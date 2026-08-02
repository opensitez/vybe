// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_nested_string_arrays_joined_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

var flat=new[]{new[]{"a","b"},new[]{"c"}}.SelectMany(x=>x);
foreach(var s in flat) Console.WriteLine(s);
