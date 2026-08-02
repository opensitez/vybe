// vybe-test: csharp/collections_advanced/dict_trygetvalue
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new Dictionary<string, int> { { "age", 30 } };
int value;
if (dict.TryGetValue("age", out value)) {
    __Check((value).ToString(), "30");
}
if (!dict.TryGetValue("name", out value)) {
    __Check(("not found").ToString(), "not found");
}
