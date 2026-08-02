// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_receive_argument_from_first_part
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker { partial void OnRun(int value); public void Run() { OnRun(5); } } partial class Worker { partial void OnRun(int value) { System.__Check((value * 2).ToString(), "10"); } } new Worker().Run();
