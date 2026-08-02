// vybe-test: csharp/collections/list_sort
// origin: languages/csharp/tests/csharp/test_collections.rs

var list = new List<int>();
        list.Add(3);
        list.Add(1);
        list.Add(2);
        list.Sort();
        foreach (var x in list) { Console.WriteLine(x); }
