// vybe-test: csharp/control_flow/break_in_loop
// origin: languages/csharp/tests/csharp/test_control_flow.rs

var result = 0;
        for (var i = 0; i < 100; i++) {
            if (i == 5) break;
            result = result + 1;
        }
        Console.WriteLine(result);
