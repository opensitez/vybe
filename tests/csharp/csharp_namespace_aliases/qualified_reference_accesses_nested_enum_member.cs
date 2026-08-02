// vybe-test: csharp/csharp_namespace_aliases/qualified_reference_accesses_nested_enum_member
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public class Job { public enum State { Idle, Done } } } __Check((Demo.Job.State.Done).ToString(), "Done");
