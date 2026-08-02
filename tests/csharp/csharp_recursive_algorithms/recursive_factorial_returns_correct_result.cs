// vybe-test: csharp/csharp_recursive_algorithms/recursive_factorial_returns_correct_result
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long Fact(int n)=>n<=1?1:n*Fact(n-1);
__Check((Fact(10)).ToString(), "3628800");
