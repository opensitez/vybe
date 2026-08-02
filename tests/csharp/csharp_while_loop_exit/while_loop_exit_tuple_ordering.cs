// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
var tuple = (left: 47, right: 48); __Check((tuple.left < tuple.right).ToString(), "True");
