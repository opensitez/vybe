// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_interface_is_implemented_by_qualified_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public interface IRun { string Run(); } public class Worker : IRun { public string Run() { return "done"; } } } Demo.IRun worker = new Demo.Worker(); __Check((worker.Run()).ToString(), "done");
