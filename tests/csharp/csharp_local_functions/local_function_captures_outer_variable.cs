// vybe-test: csharp/csharp_local_functions/local_function_captures_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int multiplier=3;
int Mul(int n){
    int Scaled(int x)=>x*multiplier;
    return Scaled(n);
}
__Check((Mul(7)).ToString(), "21");
