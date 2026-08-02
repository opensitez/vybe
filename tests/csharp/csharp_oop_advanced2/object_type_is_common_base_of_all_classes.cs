// vybe-test: csharp/csharp_oop_advanced2/object_type_is_common_base_of_all_classes
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object x=42; object y="hi"; object z=new int[]{};
__Check((x is object).ToString(), "True");
__Check((y is object).ToString(), "True");
__Check((z is object).ToString(), "True");
