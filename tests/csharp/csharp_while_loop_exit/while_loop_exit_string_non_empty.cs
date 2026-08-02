// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
string feature = "while_loop_exit"; __Check((feature.Length > 0).ToString(), "True");
