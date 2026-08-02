// vybe-test: csharp/csharp_primary_constructors/primary_constructor_bool_param_controls_branch
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Gate(bool open) { public string State() => open ? "open" : "closed"; }
__Check((new Gate(true).State()).ToString(), "open");
