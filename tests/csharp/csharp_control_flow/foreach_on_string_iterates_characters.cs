// vybe-test: csharp/csharp_control_flow/foreach_on_string_iterates_characters
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

string text = "ab";
int count = 0;
foreach (var ch in text) count++;
Console.WriteLine(count);
