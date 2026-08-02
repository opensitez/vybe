// vybe-test: csharp/csharp_exceptions_flow/multi_catch_picks_first_matching_clause
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{throw new System.ArgumentNullException("x");}
catch(System.ArgumentOutOfRangeException){r="range";}
catch(System.ArgumentNullException){r="null";}
catch(System.Exception){r="general";}
__Check((r).ToString(), "null");
