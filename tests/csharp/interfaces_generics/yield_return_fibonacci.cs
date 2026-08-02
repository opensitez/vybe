// vybe-test: csharp/interfaces_generics/yield_return_fibonacci
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

class Fib {
    public static IEnumerable<int> Sequence(int count) {
        int a = 0, b = 1;
        for (int i = 0; i < count; i++) {
            yield return a;
            int temp = a + b;
            a = b;
            b = temp;
        }
    }
}
foreach (var n in Fib.Sequence(8)) Console.WriteLine(n);
