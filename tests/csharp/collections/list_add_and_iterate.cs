// vybe-test: csharp/collections/list_add_and_iterate
// origin: languages/csharp/tests/csharp/test_collections.rs

var list = new List<string>();
        list.Add("a");
        list.Add("b");
        list.Add("c");
        foreach (var item in list) {
            Console.WriteLine(item);
        }
