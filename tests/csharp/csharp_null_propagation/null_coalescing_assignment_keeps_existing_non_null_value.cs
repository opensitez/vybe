// vybe-test: csharp/csharp_null_propagation/null_coalescing_assignment_keeps_existing_non_null_value
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string value = "keep"; value ??= "set"; __Check((value).ToString(), "keep");
