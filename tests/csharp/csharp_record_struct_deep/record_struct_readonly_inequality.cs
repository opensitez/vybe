// vybe-test: csharp/csharp_record_struct_deep/record_struct_readonly_inequality
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly record struct Pair(int A,int B); __Check((new Pair(1,2)!=new Pair(2,1)).ToString(), "True");
