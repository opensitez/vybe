// vybe-test: csharp/csharp_pattern_list/is_list_string_array_exact_sequence
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] words=new[]{"a","b"}; __Check((words is ["a","b"]).ToString(), "True");
