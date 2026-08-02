// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRun { void Run(); }
class Job : IRun {
    public void Run() { Trace.Show(); }
}
class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Run");
}
IRun job = new Job(); job.Run();
