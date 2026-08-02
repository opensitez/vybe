// vybe-test: csharp/csharp_exceptions_flow/catch_when_filter_skips_non_matching_predicate
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="unhandled";
try{throw new System.Exception("skip");}
catch(System.Exception ex) when(ex.Message=="match"){r="matched";}
catch(System.Exception){r="caught";}
__Check((r).ToString(), "caught");
