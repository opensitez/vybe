// vybe-test: csharp/control_flow/foreach_array
// origin: languages/csharp/tests/csharp/test_control_flow.rs

var sum = 0;
        foreach (var x in new int[] { 10, 20, 30 }) {
            sum = sum + x;
        }
        Console.WriteLine(sum);
