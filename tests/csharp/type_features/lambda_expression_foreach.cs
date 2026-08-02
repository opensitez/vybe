// vybe-test: csharp/type_features/lambda_expression_foreach
// origin: languages/csharp/tests/csharp/test_type_features.rs

var arr = new int[] { 1, 2, 3, 4, 5 };
        var sum = 0;
        foreach (var x in arr) { sum = sum + x; }
        Console.WriteLine(sum);
