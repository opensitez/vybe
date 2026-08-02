// vybe-test: csharp/csharp_method_overloading/overload_with_different_argument_count_dispatches_correctly
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Desc(int a)=>"one";
string Desc(int a,int b)=>"two";
__Check((Desc(1)).ToString(), "one"); __Check((Desc(1,2)).ToString(), "two");
