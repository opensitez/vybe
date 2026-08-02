// vybe-test: csharp/csharp_rectangular_array_traversal/rectangular_array_constructor_sets_both_dimensions
// origin: languages/csharp/tests/csharp/test_csharp_rectangular_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var grid = new int[2, 3];
__Check((grid.GetLength(0)).ToString(), "2");
__Check((grid.GetLength(1)).ToString(), "3");
