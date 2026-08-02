// vybe-test: csharp/csharp_arrays_advanced/array_for_loop
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

var arr = new[] { 1, 2, 3, 4, 5 };
int sum = 0;
for (int i = 0; i < arr.Length; i++) {
    sum += arr[i];
}
Console.WriteLine(sum);
