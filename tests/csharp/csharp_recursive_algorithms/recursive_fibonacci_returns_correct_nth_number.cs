// vybe-test: csharp/csharp_recursive_algorithms/recursive_fibonacci_returns_correct_nth_number
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Fib(int n)=>n<=1?n:Fib(n-1)+Fib(n-2);
__Check((Fib(8)).ToString(), "21");
