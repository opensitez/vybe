// vybe-test: csharp/modern_features/tuple_return_from_method
// origin: languages/csharp/tests/csharp/test_modern_features.rs

class MathOps {
    public static (int Min, int Max) MinMax(int[] arr) {
        int min = arr[0], max = arr[0];
        foreach (var x in arr) {
            if (x < min) min = x;
            if (x > max) max = x;
        }
        return (min, max);
    }
}
var result = MathOps.MinMax(new int[] { 3, 1, 4, 1, 5, 9 });
Console.WriteLine(result.Min);
Console.WriteLine(result.Max);
