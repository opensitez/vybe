// vybe-test: csharp/csharp_closures/closure_sees_mutation_of_captured_variable
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 0;
System.Func<int> read = () => x;
x = 99;
__Check((read()).ToString(), "99");
