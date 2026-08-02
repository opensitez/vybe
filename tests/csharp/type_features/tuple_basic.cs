// vybe-test: csharp/type_features/tuple_basic
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (1, "hello", true);
        __Check((t[0]).ToString(), "1");
        __Check((t[1]).ToString(), "hello");
        __Check((t[2]).ToString(), "True");
