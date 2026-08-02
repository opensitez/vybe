// vybe-test: csharp/csharp_object_initializers/array_initializer_infers_element_type
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr=new[]{1,2,3};
__Check((arr.GetType().IsArray).ToString(), "True"); __Check((arr.Length).ToString(), "3");
