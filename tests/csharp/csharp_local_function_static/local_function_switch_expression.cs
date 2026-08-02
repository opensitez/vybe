// vybe-test: csharp/csharp_local_function_static/local_function_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(int n){string L(int x)=>x switch{1=>"one",2=>"two",_=>"other"}; return L(n);} __Check((Label(2)).ToString(), "two");
