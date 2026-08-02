// vybe-test: csharp/csharp_linq_projections/select_with_index_provides_position
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{"a","b","c"}.Select((x,i) => $"{i}:{x}");
foreach(var s in result) Console.WriteLine(s);
