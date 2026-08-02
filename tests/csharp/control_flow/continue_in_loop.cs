// vybe-test: csharp/control_flow/continue_in_loop
// origin: languages/csharp/tests/csharp/test_control_flow.rs

var sum = 0;
        for (var i = 0; i < 10; i++) {
            if (i % 2 != 0) continue;
            sum = sum + i;
        }
        Console.WriteLine(sum);
