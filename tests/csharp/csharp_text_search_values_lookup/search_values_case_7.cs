// vybe-test: csharp/csharp_text_search_values_lookup/search_values_case_7

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

var sv = System.Buffers.SearchValues.Create(new char[] { 'a', 'e', 'i', 'o', 'u' });
ReadOnlySpan<char> span = "test_item_7".AsSpan();
int idx = span.IndexOfAny(sv);
__P((idx >= 0).ToString());
__Check("True");
