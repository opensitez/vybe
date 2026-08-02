// vybe-test: csharp/interfaces_generics/yield_return_basic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

class Numbers {
    public static IEnumerable<int> OneToFive() {
        yield return 1;
        yield return 2;
        yield return 3;
        yield return 4;
        yield return 5;
    }
}
foreach (var n in Numbers.OneToFive()) Console.WriteLine(n);
