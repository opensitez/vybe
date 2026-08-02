// vybe-test: csharp/csharp_expression_bodied_members/expr_property_struct_get_only
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pair { public int A, B; public int Sum => A + B; }
var p = new Pair { A = 2, B = 5 }; __Check((p.Sum).ToString(), "7");
