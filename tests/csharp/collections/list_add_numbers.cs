// vybe-test: csharp/collections/list_add_numbers
// origin: languages/csharp/tests/csharp/test_collections.rs

var list = new List<int>();
        list.Add(10);
        list.Add(20);
        list.Add(30);
        var sum = 0;
        foreach (var x in list) { sum = sum + x; }
        Console.WriteLine(sum);
