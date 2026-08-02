// vybe-test: csharp/oop_advanced/indexer_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sentence {
    string[] words;
    public Sentence(string[] w) { words = w; }
    public string this[int index] {
        get { return words[index]; }
        set { words[index] = value; }
    }
}
var s = new Sentence(new string[] { "hello", "world" });
__Check((s[0]).ToString(), "hello");
__Check((s[1]).ToString(), "world");
s[1] = "C#";
__Check((s[1]).ToString(), "C#");
