// vybe-test: csharp/type_features/generic_method_call
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int>();
        list.Add(1);
        list.Add(2);
        list.Add(3);
        __Check((list.Count).ToString(), "3");
