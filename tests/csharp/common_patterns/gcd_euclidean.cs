// vybe-test: csharp/common_patterns/gcd_euclidean
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

class Algorithms {
    public static int GCD(int a, int b) {
        while (b != 0) { int t = b; b = a % b; a = t; }
        return a;
    }
}
Console.WriteLine(Algorithms.GCD(48, 18));
Console.WriteLine(Algorithms.GCD(100, 75));
