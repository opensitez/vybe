// vybe-test: csharp/csharp_text_stringbuilder/string_builder_capacity_grows_automatically
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

var sb=new System.Text.StringBuilder(4);
for(int i=0;i<100;i++) sb.Append('x');
Console.WriteLine(sb.Length);
