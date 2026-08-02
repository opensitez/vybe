// vybe-test: csharp/csharp_oop_polymorphism/direct_cast_throws_invalid_cast_for_unrelated_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{object o="hello"; int n=(int)o;}
catch(System.InvalidCastException){r="bad cast";}
__Check((r).ToString(), "bad cast");
