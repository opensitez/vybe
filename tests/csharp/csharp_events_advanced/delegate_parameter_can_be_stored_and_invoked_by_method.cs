// vybe-test: csharp/csharp_events_advanced/delegate_parameter_can_be_stored_and_invoked_by_method
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Runner { public static void Run(Action action) { action(); } } Runner.Run(() => __Check(("go").ToString(), "go"));
