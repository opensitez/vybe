// vybe-test: csharp/oop_advanced/indexer_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((s[0]).ToString());
__P((s[1]).ToString());
s[1] = "C#";
__P((s[1]).ToString());
__Check("hello\nworld\nC#");
