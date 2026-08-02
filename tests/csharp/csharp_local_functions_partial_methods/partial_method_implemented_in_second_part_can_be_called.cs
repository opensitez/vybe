// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_implemented_in_second_part_can_be_called
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker { partial void OnRun(); public void Run() { OnRun(); } } partial class Worker { partial void OnRun() { System.__Check(("ran").ToString(), "ran"); } } new Worker().Run();
