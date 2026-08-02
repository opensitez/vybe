// vybe-test: csharp/csharp_delegates_advanced/lambda_closed_over_mutable_list_builds_result
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var log=new System.Collections.Generic.List<string>();
System.Action<string> record=msg=>log.Add(msg);
record("a"); record("b"); record("c");
__Check((string.Join(",",log)).ToString(), "a,b,c");
