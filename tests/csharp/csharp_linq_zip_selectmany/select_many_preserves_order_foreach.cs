// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_preserves_order_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

var flat=new[]{new[]{1,2},new[]{3,4}}.SelectMany(x=>x);
foreach(var n in flat) Console.WriteLine(n);
