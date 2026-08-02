// vybe-test: csharp/csharp_record_struct_deep/record_struct_float_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Sample(float R); __Check((new Sample(1.5f)==new Sample(1.5f)).ToString(), "True");
