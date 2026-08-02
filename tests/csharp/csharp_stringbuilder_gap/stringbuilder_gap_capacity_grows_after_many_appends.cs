// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_capacity_grows_after_many_appends
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

var sb=new System.Text.StringBuilder(4); for(int i=0;i<50;i++) sb.Append('q'); Console.WriteLine(sb.Capacity>=50);
