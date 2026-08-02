// vybe-test: csharp/csharp_modern/params_array
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

int Sum(params int[] nums) {
    int total = 0;
    foreach (var n in nums) total += n;
    return total;
}
Console.WriteLine(Sum(1, 2, 3, 4));
