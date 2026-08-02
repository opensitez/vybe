// vybe-test: csharp/csharp_hashcode/hashcode_add_produces_same_as_combine_for_two_values
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var hc=new System.HashCode();
hc.Add(1); hc.Add(2);
int h1=hc.ToHashCode();
int h2=System.HashCode.Combine(1,2);
__Check((h1==h2).ToString(), "True");
