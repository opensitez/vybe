// vybe-test: csharp/csharp_records_advanced/record_with_additional_method_can_compute_value
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Counter(int Value) { public int Double() { return Value * 2; } } __Check((new Counter(6).Double()).ToString(), "12");
