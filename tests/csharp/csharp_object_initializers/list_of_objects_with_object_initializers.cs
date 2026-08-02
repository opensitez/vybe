// vybe-test: csharp/csharp_object_initializers/list_of_objects_with_object_initializers
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item{public int Id;}
var items=new System.Collections.Generic.List<Item>{new Item{Id=1},new Item{Id=2}};
__Check((items[1].Id).ToString(), "2");
