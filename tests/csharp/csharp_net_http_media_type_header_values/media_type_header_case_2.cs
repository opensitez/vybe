// vybe-test: csharp/csharp_net_http_media_type_header_values/media_type_header_case_2

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

var header = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json; charset=utf-8");
__P(header.MediaType);
__P(header.CharSet);
__Check("application/json\nutf-8");
