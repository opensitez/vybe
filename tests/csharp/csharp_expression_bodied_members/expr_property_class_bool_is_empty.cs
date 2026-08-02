// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_bool_is_empty
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag { public string? Data; public bool IsEmpty => Data == null || Data.Length == 0; }
__Check((new Bag { Data = "" }.IsEmpty).ToString(), "True"); __Check((new Bag { Data = "x" }.IsEmpty).ToString(), "False");
