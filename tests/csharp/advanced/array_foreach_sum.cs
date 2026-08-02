// vybe-test: csharp/advanced/array_foreach_sum
// origin: languages/csharp/tests/csharp/test_advanced.rs

var nums = new int[] { 10, 20, 30, 40 };
        var total = 0;
        foreach (var n in nums) { total = total + n; }
        Console.WriteLine(total);
