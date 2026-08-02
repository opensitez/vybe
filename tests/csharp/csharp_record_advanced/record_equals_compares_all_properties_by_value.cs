// vybe-test: csharp/csharp_record_advanced/record_equals_compares_all_properties_by_value
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(int A,int B);
var p1=new Pair(1,2); var p2=new Pair(1,2); var p3=new Pair(1,3);
__Check((p1==p2).ToString(), "True");
__Check((p1==p3).ToString(), "False");
