// vybe-test: csharp/collections_advanced/jagged_array
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[][] jagged = new int[3][];
jagged[0] = new int[] { 1, 2 };
jagged[1] = new int[] { 3, 4, 5 };
jagged[2] = new int[] { 6 };
__Check((jagged[0].Length).ToString(), "2");
__Check((jagged[1].Length).ToString(), "3");
__Check((jagged[1][2]).ToString(), "5");
