// vybe-test: csharp/csharp_deferred_execution/first_throws_if_sequence_is_empty
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{System.Array.Empty<int>().First();}
catch(System.InvalidOperationException){r="empty";}
__Check((r).ToString(), "empty");
