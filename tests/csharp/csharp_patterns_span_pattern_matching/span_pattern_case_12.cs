// vybe-test: csharp/csharp_patterns_span_pattern_matching/span_pattern_case_12

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

ReadOnlySpan<char> span = "item_12".AsSpan();
bool isItem = span is ['i', 't', 'e', 'm', '_', ..];
__P(isItem.ToString());
__Check("True");
