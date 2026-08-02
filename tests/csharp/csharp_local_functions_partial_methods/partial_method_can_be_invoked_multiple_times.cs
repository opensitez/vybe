// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_be_invoked_multiple_times
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

partial class Worker { partial void OnRun(); public void RunTwice() { OnRun(); OnRun(); } } partial class Worker { partial void OnRun() { System.Console.WriteLine("tick"); } } new Worker().RunTwice();
