// vybe-test: csharp/threading_dotnet/fully_qualified_process_start_info_wait_for_exit_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_threading_dotnet.rs

using static __Harness;

var psi = new System.Diagnostics.ProcessStartInfo("dotnet", "--version");
__P((psi.FileName == "dotnet").ToString());
__Check("True");
public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
