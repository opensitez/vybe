// vybe-test: csharp/csharp_local_functions/local_function_declared_and_called_within_method
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Square(int n){
    int Sq(int x)=>x*x;
    return Sq(n);
}
__Check((Square(5)).ToString(), "25");
