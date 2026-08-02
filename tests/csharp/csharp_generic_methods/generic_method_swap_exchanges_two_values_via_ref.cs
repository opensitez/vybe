// vybe-test: csharp/csharp_generic_methods/generic_method_swap_exchanges_two_values_via_ref
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Swap<T>(ref T a,ref T b){T tmp=a;a=b;b=tmp;}
int x=1,y=2; Swap(ref x,ref y);
__Check((x).ToString(), "2"); __Check((y).ToString(), "1");
