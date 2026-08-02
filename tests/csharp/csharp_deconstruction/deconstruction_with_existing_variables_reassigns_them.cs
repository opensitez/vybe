// vybe-test: csharp/csharp_deconstruction/deconstruction_with_existing_variables_reassigns_them
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int first = 0;
int second = 0;
(first, second) = (7, 9);
__Check((first).ToString(), "7");
__Check((second).ToString(), "9");
