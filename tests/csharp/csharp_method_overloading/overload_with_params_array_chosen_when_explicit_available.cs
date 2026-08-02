// vybe-test: csharp/csharp_method_overloading/overload_with_params_array_chosen_when_explicit_available
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Sum(int a,int b)=>"two";
string Sum(params int[] ns)=>"params";
__Check((Sum(1,2)).ToString(), "two");
__Check((Sum(1,2,3)).ToString(), "params");
