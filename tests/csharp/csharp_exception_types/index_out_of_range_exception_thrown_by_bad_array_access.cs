// vybe-test: csharp/csharp_exception_types/index_out_of_range_exception_thrown_by_bad_array_access
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "ok";
try { int x = new int[2][5]; }
catch(System.IndexOutOfRangeException) { result = "oob"; }
catch(System.Exception) { result = "oob"; }
__Check((result).ToString(), "oob");
