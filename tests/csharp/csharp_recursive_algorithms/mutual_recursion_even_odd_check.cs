// vybe-test: csharp/csharp_recursive_algorithms/mutual_recursion_even_odd_check
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool IsEven(int n){if(n==0)return true; return IsOdd(n-1);}
bool IsOdd(int n){if(n==0)return false; return IsEven(n-1);}
__Check((IsEven(4)).ToString(), "True"); __Check((IsOdd(3)).ToString(), "True");
