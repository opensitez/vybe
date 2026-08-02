// vybe-test: csharp/csharp_constructor_chains/field_initializer_runs_before_constructor_body_reads_value
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { string name = "init"; public Box() { __Check((name).ToString(), "init"); } } new Box();
