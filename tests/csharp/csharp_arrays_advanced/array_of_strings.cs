// vybe-test: csharp/csharp_arrays_advanced/array_of_strings
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words = new[] { "hello", "world" };
__Check((words[0] + " " + words[1]).ToString(), "hello world");
__Check((words.Length).ToString(), "2");
