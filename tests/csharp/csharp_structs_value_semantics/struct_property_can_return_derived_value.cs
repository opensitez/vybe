// vybe-test: csharp/csharp_structs_value_semantics/struct_property_can_return_derived_value
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Rect { public int W { get; set; } public int H { get; set; } public int Area => W * H; } var rect = new Rect { W = 3, H = 5 }; __Check((rect.Area).ToString(), "15");
