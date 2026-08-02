// vybe-test: csharp/csharp_control_flow/foreach_array
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int[] arr = {10, 20, 30};
int sum = 0;
foreach (var x in arr) {
    sum += x;
}
Console.WriteLine(sum);
