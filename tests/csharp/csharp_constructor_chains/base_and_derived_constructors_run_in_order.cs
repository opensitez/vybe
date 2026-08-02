// vybe-test: csharp/csharp_constructor_chains/base_and_derived_constructors_run_in_order
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public Base() { __Check(("base").ToString(), "base"); } } class Child : Base { public Child() { __Check(("child").ToString(), "child"); } } new Child();
