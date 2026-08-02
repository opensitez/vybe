// vybe-test: csharp/csharp_control_flow/switch_with_break
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int day = 3;
switch (day) {
    case 1: Console.WriteLine("Mon"); break;
    case 2: Console.WriteLine("Tue"); break;
    case 3: Console.WriteLine("Wed"); break;
    default: Console.WriteLine("Other"); break;
}
