// vybe-test: csharp/csharp_environment_system/environment_get_environment_variable_returns_null_for_unknown
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var v=System.Environment.GetEnvironmentVariable("__VYBE_NOSUCH_VAR__123");
__Check((v==null).ToString(), "True");
