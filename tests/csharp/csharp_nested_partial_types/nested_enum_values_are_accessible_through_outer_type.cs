// vybe-test: csharp/csharp_nested_partial_types/nested_enum_values_are_accessible_through_outer_type
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Job {
    public enum State { Pending, Running, Done }
}
__Check((Job.State.Pending).ToString(), "Pending");
__Check(((int)Job.State.Done).ToString(), "2");
