// vybe-test: csharp/csharp_arrays_advanced/array_passed_to_method
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

class Utils {
    public static int Sum(int[] arr) {
        int total = 0;
        foreach (var x in arr) total += x;
        return total;
    }
}
var nums = new[] { 1, 2, 3, 4 };
Console.WriteLine(Utils.Sum(nums));
