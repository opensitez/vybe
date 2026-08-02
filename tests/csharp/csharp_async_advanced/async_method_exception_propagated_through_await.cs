// vybe-test: csharp/csharp_async_advanced/async_method_exception_propagated_through_await
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Fail()=>throw new System.Exception("async fail");
string msg="";
try{await Fail();}catch(System.Exception ex){msg=ex.Message;}
__Check((msg).ToString(), "async fail");
