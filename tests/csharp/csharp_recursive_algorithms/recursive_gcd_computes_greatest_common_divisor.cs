// vybe-test: csharp/csharp_recursive_algorithms/recursive_gcd_computes_greatest_common_divisor
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Gcd(int a,int b)=>b==0?a:Gcd(b,a%b);
__Check((Gcd(48,18)).ToString(), "6");
