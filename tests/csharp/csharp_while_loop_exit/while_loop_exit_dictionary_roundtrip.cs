// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// while_loop_exit
var map = new System.Collections.Generic.Dictionary<int, int>(); map[47] = 48; __Check((map.ContainsKey(47) && map[47] == 48).ToString(), "True");
