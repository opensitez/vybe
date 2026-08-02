// vybe-test: csharp/csharp_using_declarations/using_var_list_mutation_visible_before_dispose
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var list=new System.Collections.Generic.List<int>(); list.Add(4); __Check((list[0]).ToString(), "4");
