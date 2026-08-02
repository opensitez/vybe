// vybe-test: csharp/csharp_exceptions_hierarchy/custom_exception_stores_custom_message
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class AppEx:System.Exception{public AppEx(string m):base(m){}}
string r="";
try{throw new AppEx("fail");}
catch(AppEx ex){r=ex.Message;}
__Check((r).ToString(), "fail");
