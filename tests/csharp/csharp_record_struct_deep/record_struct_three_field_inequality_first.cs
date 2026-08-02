// vybe-test: csharp/csharp_record_struct_deep/record_struct_three_field_inequality_first
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Rgb(byte R,byte G,byte B); __Check((new Rgb(10,20,30)==new Rgb(11,20,30)).ToString(), "False");
