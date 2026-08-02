// vybe-test: csharp/csharp_dynamic/dynamic_expando_object_accepts_arbitrary_properties
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

dynamic obj=new System.Dynamic.ExpandoObject();
obj.Name="Alice";
obj.Age=30;
__Check((obj.Name).ToString(), "Alice"); __Check((obj.Age).ToString(), "30");
