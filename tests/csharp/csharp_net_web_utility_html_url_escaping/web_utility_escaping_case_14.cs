// vybe-test: csharp/csharp_net_web_utility_html_url_escaping/web_utility_escaping_case_14

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

string encoded = System.Net.WebUtility.HtmlEncode("<div id='14'>");
string decoded = System.Net.WebUtility.HtmlDecode(encoded);
__P(decoded);
__Check("<div id='14'>");
