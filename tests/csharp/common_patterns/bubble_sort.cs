// vybe-test: csharp/common_patterns/bubble_sort
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

int[] arr = { 5, 3, 8, 1, 2 };
for (int i = 0; i < arr.Length; i++) {
    for (int j = 0; j < arr.Length - 1 - i; j++) {
        if (arr[j] > arr[j + 1]) {
            int tmp = arr[j];
            arr[j] = arr[j + 1];
            arr[j + 1] = tmp;
        }
    }
}
Console.WriteLine(string.Join(",", arr));
