// vybe-test: csharp/csharp_switch_expression_core/switch_expr_return_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Word(int n)=>n switch{1=>"a",2=>"b",_=>"z"}; __Check((Word(2)).ToString(), "b");
