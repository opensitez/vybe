// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_nested_string_field
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item{public string Label="""tag""";} var item=new Item(); string text=$"""label={item.Label}"""; __Check((text.Contains("tag")).ToString(), "True");
