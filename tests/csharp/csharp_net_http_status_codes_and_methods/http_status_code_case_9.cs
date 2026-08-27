// vybe-test: csharp/csharp_net_http_status_codes_and_methods/http_status_code_case_9

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

var status = System.Net.HttpStatusCode.OK;
var method = System.Net.Http.HttpMethod.Get;
__P(((int)status).ToString());
__P(method.Method);
__Check("200\nGET");
