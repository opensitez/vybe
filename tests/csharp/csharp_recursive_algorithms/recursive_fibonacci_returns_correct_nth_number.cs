// vybe-test: csharp/csharp_recursive_algorithms/recursive_fibonacci_returns_correct_nth_number
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int Fib(int n)=>n<=1?n:Fib(n-1)+Fib(n-2);
__P((Fib(8)).ToString());
__Check("21");
