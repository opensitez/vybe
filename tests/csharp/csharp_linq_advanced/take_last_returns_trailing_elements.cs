// vybe-test: csharp/csharp_linq_advanced/take_last_returns_trailing_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

var result=new[]{1,2,3,4,5}.TakeLast(2);
foreach(var n in result) Console.WriteLine(n);
