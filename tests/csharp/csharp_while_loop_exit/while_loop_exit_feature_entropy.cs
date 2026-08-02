// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
string feature = "while_loop_exit:47"; __Check((feature.Length >= 1).ToString(), "True");
