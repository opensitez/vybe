// vybe-test: csharp/csharp_environment_system/environment_newline_is_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Environment.NewLine.Length>0).ToString(), "True");
