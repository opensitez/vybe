// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_int_cast_from_outside
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Job{public enum State{Idle=0,Busy=1}} __Check(((int)Job.State.Busy).ToString(), "1");
