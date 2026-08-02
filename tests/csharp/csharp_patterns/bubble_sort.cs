// vybe-test: csharp/csharp_patterns/bubble_sort
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

var arr = new[] { 5, 3, 8, 1, 2 };
for (int i = 0; i < arr.Length - 1; i++) {
    for (int j = 0; j < arr.Length - 1 - i; j++) {
        if (arr[j] > arr[j + 1]) {
            int temp = arr[j];
            arr[j] = arr[j + 1];
            arr[j + 1] = temp;
        }
    }
}
foreach (var x in arr) Console.WriteLine(x);
