// vybe-test: csharp/csharp_control_flow/nested_loops
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int count = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 4; j++) {
        count++;
    }
}
Console.WriteLine(count);
