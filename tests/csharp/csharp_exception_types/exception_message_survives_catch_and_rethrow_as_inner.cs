// vybe-test: csharp/csharp_exception_types/exception_message_survives_catch_and_rethrow_as_inner
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string msg = "";
try {
    try { throw new System.Exception("root"); }
    catch(System.Exception e) { throw new System.Exception("wrap", e); }
} catch(System.Exception outer) { msg = outer.InnerException.Message; }
__Check((msg).ToString(), "root");
