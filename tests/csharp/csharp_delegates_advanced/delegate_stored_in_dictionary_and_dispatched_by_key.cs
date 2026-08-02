// vybe-test: csharp/csharp_delegates_advanced/delegate_stored_in_dictionary_and_dispatched_by_key
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ops=new System.Collections.Generic.Dictionary<string,System.Func<int,int,int>>{
    {"add",(a,b)=>a+b},
    {"mul",(a,b)=>a*b}
};
__Check((ops["add"](3,4)).ToString(), "7");
__Check((ops["mul"](3,4)).ToString(), "12");
