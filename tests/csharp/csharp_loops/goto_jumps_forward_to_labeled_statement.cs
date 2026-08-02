// vybe-test: csharp/csharp_loops/goto_jumps_forward_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 0;
goto done;
x = 99;
done:
__Check((x).ToString(), "0");
