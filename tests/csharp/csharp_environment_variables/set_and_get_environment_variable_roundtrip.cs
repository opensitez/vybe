// vybe-test: csharp/csharp_environment_variables/set_and_get_environment_variable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Environment.SetEnvironmentVariable("VYBE_TEST_KEY","hello");
__Check((System.Environment.GetEnvironmentVariable("VYBE_TEST_KEY")).ToString(), "hello");
