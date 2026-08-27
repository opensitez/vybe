// vybe-test: csharp/csharp_environment_variables/deleting_environment_variable_makes_it_null
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables.rs

using static __Harness;

System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY","x");
System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY",null);
__P((System.Environment.GetEnvironmentVariable("VYBE_DEL_KEY")==null).ToString());
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
