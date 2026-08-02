// vybe-test: csharp/modern_features/range_operator_basic
// origin: languages/csharp/tests/csharp/test_modern_features.rs

int[] nums = { 1, 2, 3, 4, 5 };
int[] slice = nums[1..4];
foreach (var n in slice) Console.WriteLine(n);
