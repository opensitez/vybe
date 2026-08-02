// vybe-test: csharp/csharp_environment_system/appdomain_current_domain_not_null
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.AppDomain.CurrentDomain!=null).ToString(), "True");
