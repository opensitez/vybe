// vybe-test: csharp/csharp_control_flow/break_in_loop
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

for (int i = 0; i < 100; i++) {
    if (i >= 3) break;
    Console.WriteLine(i);
}
