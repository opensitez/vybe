// vybe-test: csharp/csharp_linq_zip_selectmany/zip_enumerate_products_in_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

var z=new[]{2,3}.Zip(new[]{4,5},(a,b)=>a*b);
foreach(var n in z) Console.WriteLine(n);
