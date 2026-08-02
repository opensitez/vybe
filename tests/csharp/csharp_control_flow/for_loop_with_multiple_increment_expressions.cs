// vybe-test: csharp/csharp_control_flow/for_loop_with_multiple_increment_expressions
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int sum = 0;
for (int i = 0, j = 10; i < 3; i++, j--) {
    sum += i + j;
}
Console.WriteLine(sum);
