// vybe-test: csharp/csharp_object_initializers/dictionary_initializer_uses_key_value_syntax
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=new System.Collections.Generic.Dictionary<string,int>{{"a",1},{"b",2}};
__Check((d["b"]).ToString(), "2");
