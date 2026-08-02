// vybe-test: csharp/csharp_reflection/property_info_set_value_writes_property_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public int Id {get;set;} }
var item = new Item();
var prop = typeof(Item).GetProperty("Id");
prop.SetValue(item, 99);
__Check((item.Id).ToString(), "99");
