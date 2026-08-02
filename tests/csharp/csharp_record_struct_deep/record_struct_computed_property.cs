// vybe-test: csharp/csharp_record_struct_deep/record_struct_computed_property
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Rect(int W,int H){public int Area=>W*H;} __Check((new Rect(3,4).Area).ToString(), "12");
