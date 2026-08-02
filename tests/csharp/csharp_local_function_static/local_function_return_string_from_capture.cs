// vybe-test: csharp/csharp_local_function_static/local_function_return_string_from_capture
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string suffix="!"; string Exclaim(string text){string E(string t)=>t+suffix; return E(text);} __Check((Exclaim("hi")).ToString(), "hi!");
