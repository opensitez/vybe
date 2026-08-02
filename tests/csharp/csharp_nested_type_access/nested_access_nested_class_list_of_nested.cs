// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_list_of_nested
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag{public class Item{public int Id;} public System.Collections.Generic.List<Item> All(){var list=new System.Collections.Generic.List<Item>(); list.Add(new Item{Id=1}); return list;}} __Check((new Bag().All()[0].Id).ToString(), "1");
