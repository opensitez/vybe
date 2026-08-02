// vybe-test: csharp/common_patterns/as_operator
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj = "hello";
string s = obj as string;
__Check((s != null ? s : "null").ToString(), "hello");
int? i = obj as int?;
__Check((i != null ? i.ToString() : "null").ToString(), "null");
