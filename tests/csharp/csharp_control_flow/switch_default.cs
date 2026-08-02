// vybe-test: csharp/csharp_control_flow/switch_default
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int x = 99;
switch (x) {
    case 1: Console.WriteLine("one"); break;
    default: Console.WriteLine("other"); break;
}
