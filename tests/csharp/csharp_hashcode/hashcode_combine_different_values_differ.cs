// vybe-test: csharp/csharp_hashcode/hashcode_combine_different_values_differ
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int h1=System.HashCode.Combine(1,2);
int h2=System.HashCode.Combine(2,1);
__Check((h1!=h2).ToString(), "True");
