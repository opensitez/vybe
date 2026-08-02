// vybe-test: csharp/csharp_local_function_static/local_function_capture_string_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string tag="id"; string Label(int n){string L(int x)=>$"{tag}:{x}"; return L(n);} __Check((Label(9)).ToString(), "id:9");
