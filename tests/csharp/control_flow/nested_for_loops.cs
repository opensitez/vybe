// vybe-test: csharp/control_flow/nested_for_loops
// origin: languages/csharp/tests/csharp/test_control_flow.rs

var sum = 0;
        for (var i = 0; i < 3; i++) {
            for (var j = 0; j < 3; j++) {
                sum = sum + 1;
            }
        }
        Console.WriteLine(sum);
