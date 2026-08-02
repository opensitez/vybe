// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_switch_expression_body
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ISw{string Code(int n)=>n switch{1=>"one",2=>"two",_=>"many"};} class C:ISw{} __Check((new C().Code(2)).ToString(), "two");
