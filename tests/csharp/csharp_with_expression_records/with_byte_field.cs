// vybe-test: csharp/csharp_with_expression_records/with_byte_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record ByteBox(byte B); var c=(new ByteBox(1)) with{B=255}; __Check((c.B).ToString(), "255");
