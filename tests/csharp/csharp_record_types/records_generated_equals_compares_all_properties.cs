// vybe-test: csharp/csharp_record_types/records_generated_equals_compares_all_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var a = new Point(1,2); var b = new Point(1,2); var c = new Point(1,3);
__Check((a == b).ToString(), "True");
__Check((a == c).ToString(), "False");
