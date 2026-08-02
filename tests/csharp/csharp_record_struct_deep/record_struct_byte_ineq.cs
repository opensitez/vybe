// vybe-test: csharp/csharp_record_struct_deep/record_struct_byte_ineq
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct ByteVal(byte B); __Check((new ByteVal(1)==new ByteVal(2)).ToString(), "False");
