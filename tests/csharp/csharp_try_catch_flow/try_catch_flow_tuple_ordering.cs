// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
var tuple = (left: 51, right: 52); __Check((tuple.left < tuple.right).ToString(), "True");
