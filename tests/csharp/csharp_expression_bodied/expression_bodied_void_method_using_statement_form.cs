// vybe-test: csharp/csharp_expression_bodied/expression_bodied_void_method_using_statement_form
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger{public void Log(string msg)=>__Check((msg).ToString(), "hello");}
new Logger().Log("hello");
