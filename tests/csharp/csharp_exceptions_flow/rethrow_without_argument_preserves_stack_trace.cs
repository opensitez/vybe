// vybe-test: csharp/csharp_exceptions_flow/rethrow_without_argument_preserves_stack_trace
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
    catch{throw;}
}catch(System.Exception ex){r=ex.Message;}
__Check((r).ToString(), "orig");
