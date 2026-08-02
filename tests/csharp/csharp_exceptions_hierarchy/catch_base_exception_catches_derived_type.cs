// vybe-test: csharp/csharp_exceptions_hierarchy/catch_base_exception_catches_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{int[] a=new int[3]; var _=a[10];}
catch(System.Exception ex){r=ex.GetType().Name;}
__Check((r).ToString(), "IndexOutOfRangeException");
