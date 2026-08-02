// vybe-test: csharp/csharp_control_flow/continue_in_loop
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int sum = 0;
for (int i = 0; i < 10; i++) {
    if (i % 2 != 0) continue;
    sum += i;
}
Console.WriteLine(sum);
