// vybe-test: csharp/csharp_method_overload_resolution/named_arguments_can_reorder_optional_parameters_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

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

void Connect(string host, int port = 80, bool secure = false) {
    __P((host + ":" + port + ":" + secure).ToString());
}
Connect(secure: true, host: "api", port: 443);
__Check("api:443:True");
