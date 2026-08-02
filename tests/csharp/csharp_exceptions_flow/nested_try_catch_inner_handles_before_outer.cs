// vybe-test: csharp/csharp_exceptions_flow/nested_try_catch_inner_handles_before_outer
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{
    try{throw new System.Exception("inner");}
    catch(System.Exception ex){r="inner:"+ex.Message; throw new System.Exception("outer");}
}catch(System.Exception ex){r+=" outer:"+ex.Message;}
__Check((r).ToString(), "inner:inner outer:outer");
