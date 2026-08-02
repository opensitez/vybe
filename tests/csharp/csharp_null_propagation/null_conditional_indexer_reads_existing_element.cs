// vybe-test: csharp/csharp_null_propagation/null_conditional_indexer_reads_existing_element
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] values = { 3, 4, 5 }; __Check((values?[1] ?? -1).ToString(), "4");
