// vybe-test: csharp/csharp_environment_variables/current_directory_is_non_empty_string
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Environment.CurrentDirectory.Length>0).ToString(), "True");
