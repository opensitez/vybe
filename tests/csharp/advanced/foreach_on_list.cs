// vybe-test: csharp/advanced/foreach_on_list
// origin: languages/csharp/tests/csharp/test_advanced.rs

var list = new List<string>();
        list.Add("a");
        list.Add("b");
        list.Add("c");
        foreach (var item in list) {
            Console.WriteLine(item);
        }
