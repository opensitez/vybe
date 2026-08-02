// vybe-test: csharp/csharp_method_overload_resolution/named_arguments_can_reorder_optional_parameters_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Connect(string host, int port = 80, bool secure = false) {
    __Check((host + ":" + port + ":" + secure).ToString(), "api:443:True");
}
Connect(secure: true, host: "api", port: 443);
