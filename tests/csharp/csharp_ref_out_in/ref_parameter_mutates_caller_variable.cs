// vybe-test: csharp/csharp_ref_out_in/ref_parameter_mutates_caller_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Double(ref int x){x*=2;}
int n=5; Double(ref n); __Check((n).ToString(), "10");
