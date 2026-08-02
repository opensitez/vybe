// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
var values = new System.Collections.Generic.List<int> { 47, 48, 47 }; __Check((values.Count == 3).ToString(), "True");
