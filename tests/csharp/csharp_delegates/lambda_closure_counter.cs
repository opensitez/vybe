// vybe-test: csharp/csharp_delegates/lambda_closure_counter
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 0;
Action inc = () => { count++; };
inc();
inc();
inc();
__Check((count).ToString(), "3");
