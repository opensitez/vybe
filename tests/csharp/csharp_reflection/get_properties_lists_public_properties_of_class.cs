// vybe-test: csharp/csharp_reflection/get_properties_lists_public_properties_of_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public int Id {get;set;} public string Name {get;set;} }
__Check((typeof(Item).GetProperties().Length).ToString(), "2");
