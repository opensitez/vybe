// vybe-test: csharp/csharp_dynamic/dynamic_dictionary_access_via_expando
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

dynamic e=new System.Dynamic.ExpandoObject();
var dict=(System.Collections.Generic.IDictionary<string,object>)e;
dict["x"]=99;
__Check((e.x).ToString(), "99");
