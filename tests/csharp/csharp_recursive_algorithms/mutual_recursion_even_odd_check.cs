// vybe-test: csharp/csharp_recursive_algorithms/mutual_recursion_even_odd_check
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

bool IsEven(int n){if(n==0)return true; return IsOdd(n-1);}
bool IsOdd(int n){if(n==0)return false; return IsEven(n-1);}
__P((IsEven(4)).ToString()); __P((IsOdd(3)).ToString());
__Check("True\nTrue");
