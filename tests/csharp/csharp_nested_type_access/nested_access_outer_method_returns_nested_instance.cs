// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_method_returns_nested_instance
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Factory{public class Item{public string Tag="x";} public Item Build()=>new Item();} __Check((new Factory().Build().Tag).ToString(), "x");
