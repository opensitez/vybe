// vybe-test: csharp/csharp_method_overload_resolution/out_parameter_is_assigned_before_caller_observes_result
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

bool TryHalve(int input, out int half) {
    if (input % 2 != 0) {
        half = 0;
        return false;
    }
    half = input / 2;
    return true;
}
if (TryHalve(8, out var result)) {
    Console.WriteLine(result);
} else {
    Console.WriteLine("fail");
}
