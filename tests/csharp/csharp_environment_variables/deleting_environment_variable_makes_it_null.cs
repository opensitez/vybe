// vybe-test: csharp/csharp_environment_variables/deleting_environment_variable_makes_it_null
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY","x");
System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY",null);
__Check((System.Environment.GetEnvironmentVariable("VYBE_DEL_KEY")==null).ToString(), "True");
