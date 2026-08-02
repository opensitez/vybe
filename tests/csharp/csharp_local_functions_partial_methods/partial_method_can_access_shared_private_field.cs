// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_access_shared_private_field
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker { int count = 3; partial void OnRun(); public void Run() { OnRun(); } } partial class Worker { partial void OnRun() { System.__Check((count).ToString(), "3"); } } new Worker().Run();
