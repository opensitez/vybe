// vybe-test: csharp/csharp_null_propagation/null_coalescing_operator_selects_right_when_left_is_null
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string value = null; __Check((value ?? "right").ToString(), "right");
