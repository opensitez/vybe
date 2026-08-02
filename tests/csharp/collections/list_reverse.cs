// vybe-test: csharp/collections/list_reverse
// origin: languages/csharp/tests/csharp/test_collections.rs

var list = new List<int>();
        list.Add(1);
        list.Add(2);
        list.Add(3);
        list.Reverse();
        foreach (var x in list) { Console.WriteLine(x); }
