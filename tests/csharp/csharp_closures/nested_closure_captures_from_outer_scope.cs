// vybe-test: csharp/csharp_closures/nested_closure_captures_from_outer_scope
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,System.Func<int>> makeAdder = x => () => x + 1;
var add1 = makeAdder(5);
__Check((add1()).ToString(), "6");
