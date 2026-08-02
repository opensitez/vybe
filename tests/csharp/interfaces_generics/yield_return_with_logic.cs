// vybe-test: csharp/interfaces_generics/yield_return_with_logic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

class Gen {
    public static IEnumerable<int> EvenNumbers(int max) {
        for (int i = 0; i <= max; i++) {
            if (i % 2 == 0) yield return i;
        }
    }
}
foreach (var n in Gen.EvenNumbers(10)) Console.WriteLine(n);
