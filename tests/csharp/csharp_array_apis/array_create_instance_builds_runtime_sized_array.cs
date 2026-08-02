// vybe-test: csharp/csharp_array_apis/array_create_instance_builds_runtime_sized_array
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var array = System.Array.CreateInstance(typeof(int), 3); __Check((array.Length).ToString(), "3");
