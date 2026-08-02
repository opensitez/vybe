// vybe-test: csharp/csharp_pattern_list/is_list_pattern_guard_rejects_wrong_order
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

int[] data=new[]{8,4}; if(data is [var a,var b] && a<b) Console.WriteLine("ordered"); else Console.WriteLine("not");
