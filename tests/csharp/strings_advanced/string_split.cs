// vybe-test: csharp/strings_advanced/string_split
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

string csv = "a,b,c,d";
string[] parts = csv.Split(',');
foreach (var p in parts) Console.WriteLine(p);
