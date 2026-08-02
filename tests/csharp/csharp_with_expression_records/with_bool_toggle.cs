// vybe-test: csharp/csharp_with_expression_records/with_bool_toggle
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Flag(bool On); var g=(new Flag(false)) with{On=true}; __Check((g.On).ToString(), "True");
