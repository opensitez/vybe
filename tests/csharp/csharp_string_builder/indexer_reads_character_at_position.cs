// vybe-test: csharp/csharp_string_builder/indexer_reads_character_at_position
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder("xyz");
__Check((sb[1]).ToString(), "y");
