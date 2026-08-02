// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_percent_full
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tank { public int level = 75; public int capacity = 100; public int Percent => level * 100 / capacity; }
__Check((new Tank().Percent).ToString(), "75");
