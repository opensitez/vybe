// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_passes_nested_to_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Store{public class Item{public int Id;} int Inspect(Item i)=>i.Id; public int Check()=>Inspect(new Item{Id=44});} __Check((new Store().Check()).ToString(), "44");
