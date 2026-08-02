// vybe-test: csharp/common_patterns/binary_search_manual
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

int[] arr = { 1, 3, 5, 7, 9, 11, 13 };
int target = 7;
int lo = 0, hi = arr.Length - 1;
while (lo <= hi) {
    int mid = (lo + hi) / 2;
    if (arr[mid] == target) { Console.WriteLine("found at " + mid); break; }
    else if (arr[mid] < target) lo = mid + 1;
    else hi = mid - 1;
}
