// vybe-test: csharp/csharp_local_function_static/local_function_capture_bool_flag
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool enabled=true; int Gate(int n){int G(int x)=>enabled?x:-x; return G(n);} __Check((Gate(7)).ToString(), "7");
