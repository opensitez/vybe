// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_be_triggered_from_constructor
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker { partial void OnCreated(); public Worker() { OnCreated(); } } partial class Worker { partial void OnCreated() { System.__Check(("created").ToString(), "created"); } } new Worker();
