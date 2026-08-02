// vybe-test: csharp/csharp_control_flow/for_loop
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int sum = 0;
for (int i = 1; i <= 5; i++) {
    sum += i;
}
Console.WriteLine(sum);
