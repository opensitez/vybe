// vybe-test: csharp/csharp_loops/foreach_iterates_array_in_declaration_order
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int s=0; foreach(var x in new[]{3,1,4}) s+=x; Console.WriteLine(s);
