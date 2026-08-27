// vybe-test: csharp/csharp_text_composite_format_caching/composite_format_case_2

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

var cf = System.Text.CompositeFormat.Parse("Val: {0}");
string res = string.Format(System.Globalization.CultureInfo.InvariantCulture, cf, 2);
__P(res);
__Check("Val: 2");
