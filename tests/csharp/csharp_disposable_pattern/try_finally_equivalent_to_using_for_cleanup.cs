// vybe-test: csharp/csharp_disposable_pattern/try_finally_equivalent_to_using_for_cleanup
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool cleaned=false;
var f=new System.Action(()=>cleaned=true);
try{}finally{f();}
__Check((cleaned).ToString(), "True");
