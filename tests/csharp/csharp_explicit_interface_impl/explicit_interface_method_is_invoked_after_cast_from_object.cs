// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_is_invoked_after_cast_from_object
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRunner { string Run(); }
class TaskRunner : IRunner {
    string IRunner.Run() { return "done"; }
}
object item = new TaskRunner();
__Check((((IRunner)item).Run()).ToString(), "done");
