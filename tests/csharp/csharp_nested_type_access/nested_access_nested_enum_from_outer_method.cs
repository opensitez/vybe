// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_from_outer_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Job{public enum State{Idle,Busy} public State Current()=>State.Busy;} __Check((new Job().Current()).ToString(), "Busy");
