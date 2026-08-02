// vybe-test: csharp/csharp_record_types/positional_record_supports_deconstruction
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Size(int W, int H);
var s = new Size(10,20);
var (w,h) = s;
__Check((w).ToString(), "10"); __Check((h).ToString(), "20");
