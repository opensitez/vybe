// vybe-test: csharp/csharp_exceptions_hierarchy/exception_data_dictionary_stores_arbitrary_key_value
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{
    var ex=new System.Exception("test");
    ex.Data["userId"]=42;
    throw ex;
}catch(System.Exception ex){r=ex.Data["userId"].ToString();}
__Check((r).ToString(), "42");
