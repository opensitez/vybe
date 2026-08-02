// vybe-test: csharp/csharp_reflection/property_info_get_value_reads_property_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public int Id {get;set;} }
var item = new Item { Id=7 };
var prop = typeof(Item).GetProperty("Id");
__Check((prop.GetValue(item)).ToString(), "7");
