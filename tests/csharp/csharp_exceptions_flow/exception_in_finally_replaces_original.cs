// vybe-test: csharp/csharp_exceptions_flow/exception_in_finally_replaces_original
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{
    try{throw new System.Exception("orig");}
    finally{throw new System.Exception("final");}
}catch(System.Exception ex){r=ex.Message;}
__Check((r).ToString(), "final");
