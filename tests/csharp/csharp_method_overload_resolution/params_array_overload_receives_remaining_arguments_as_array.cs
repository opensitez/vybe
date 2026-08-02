// vybe-test: csharp/csharp_method_overload_resolution/params_array_overload_receives_remaining_arguments_as_array
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

int Sum(params int[] values) {
    int total = 0;
    foreach (var v in values) total += v;
    return total;
}
Console.WriteLine(Sum(1, 2, 3));
Console.WriteLine(Sum());
