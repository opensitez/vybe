// vybe-test: csharp/csharp_value_ref_semantics/null_reference_throws_null_reference_exception
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{string s=null;int len=s.Length;}
catch(System.NullReferenceException){r="null";}
__Check((r).ToString(), "null");
