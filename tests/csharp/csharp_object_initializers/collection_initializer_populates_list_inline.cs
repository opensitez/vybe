// vybe-test: csharp/csharp_object_initializers/collection_initializer_populates_list_inline
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=new System.Collections.Generic.List<int>{10,20,30};
__Check((list[1]).ToString(), "20");
