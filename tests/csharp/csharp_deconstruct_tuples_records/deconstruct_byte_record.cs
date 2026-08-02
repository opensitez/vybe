// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_byte_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record ByteVal(byte B); var (b)=new ByteVal(255); __Check((b).ToString(), "255");
