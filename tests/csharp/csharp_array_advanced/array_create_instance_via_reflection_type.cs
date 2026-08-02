// vybe-test: csharp/csharp_array_advanced/array_create_instance_via_reflection_type
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr=(int[])System.Array.CreateInstance(typeof(int),5);
arr[3]=99;
__Check((arr[3]).ToString(), "99");
