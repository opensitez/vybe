// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
var set = new System.Collections.Generic.HashSet<int>(); set.Add(51); set.Add(51); __Check((set.Count == 1).ToString(), "True");
