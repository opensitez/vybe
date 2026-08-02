// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_assigns_two_scalars
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (x, y) = (3, 4);
__Check((x).ToString(), "3");
__Check((y).ToString(), "4");
