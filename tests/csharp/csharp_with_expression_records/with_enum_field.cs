// vybe-test: csharp/csharp_with_expression_records/with_enum_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{Off,On} record State(Mode M); var t=(new State(Mode.Off)) with{M=Mode.On}; __Check((t.M).ToString(), "On");
