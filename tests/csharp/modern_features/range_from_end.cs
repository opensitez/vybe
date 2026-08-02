// vybe-test: csharp/modern_features/range_from_end
// origin: languages/csharp/tests/csharp/test_modern_features.rs

int[] nums = { 1, 2, 3, 4, 5 };
int[] last3 = nums[^3..];
foreach (var n in last3) Console.WriteLine(n);
