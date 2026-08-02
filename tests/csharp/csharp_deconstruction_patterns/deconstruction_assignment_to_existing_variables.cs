// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_assignment_to_existing_variables
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=0, y=0;
(x, y) = (5, 10);
__Check((x).ToString(), "5"); __Check((y).ToString(), "10");
