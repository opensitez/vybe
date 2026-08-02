// vybe-test: csharp/csharp_with_expression_records/with_float_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Sample(float R); var t=(new Sample(1.0f)) with{R=2.5f}; __Check((t.R).ToString(), "2.5");
