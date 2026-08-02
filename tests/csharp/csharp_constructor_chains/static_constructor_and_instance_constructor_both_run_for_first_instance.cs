// vybe-test: csharp/csharp_constructor_chains/static_constructor_and_instance_constructor_both_run_for_first_instance
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { static Box() { __Check(("static").ToString(), "static"); } public Box() { __Check(("instance").ToString(), "instance"); } } new Box();
