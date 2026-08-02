// vybe-test: csharp/csharp_reflection/field_info_get_value_reads_public_field
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Data { public int X = 3; }
var obj = new Data();
var field = typeof(Data).GetField("X");
__Check((field.GetValue(obj)).ToString(), "3");
