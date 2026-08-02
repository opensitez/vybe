// vybe-test: csharp/csharp_exceptions_hierarchy/custom_exception_with_inner_exception_chain
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer:System.Exception{public Outer(System.Exception inner):base("outer",inner){}}
string r="";
try{throw new Outer(new System.ArgumentNullException("arg"));}
catch(Outer ex){r=ex.InnerException?.GetType().Name;}
__Check((r).ToString(), "ArgumentNullException");
