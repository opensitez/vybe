// vybe-test: csharp/csharp_record_struct_deep/record_struct_min_int_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Edge(int V); __Check((new Edge(int.MinValue)==new Edge(int.MinValue)).ToString(), "True");
