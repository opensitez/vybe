// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
var set = new System.Collections.Generic.HashSet<int>(); set.Add(47); set.Add(47); __Check((set.Count == 1).ToString(), "True");
