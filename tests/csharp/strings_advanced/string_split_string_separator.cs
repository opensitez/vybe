// vybe-test: csharp/strings_advanced/string_split_string_separator
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

string s = "one::two::three";
string[] parts = s.Split("::");
foreach (var p in parts) Console.WriteLine(p);
