// vybe-test: csharp/csharp_text_rune_unicode_scalars/rune_scalar_case_20

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var r = new System.Text.Rune('A');
__P(r.IsAscii.ToString());
__P(r.Utf8SequenceLength.ToString());
__Check("True\n1");
