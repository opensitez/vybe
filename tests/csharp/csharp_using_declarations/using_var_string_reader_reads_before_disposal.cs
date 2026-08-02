// vybe-test: csharp/csharp_using_declarations/using_var_string_reader_reads_before_disposal
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var sr=new System.IO.StringReader("hi"); __Check((sr.ReadLine()).ToString(), "hi");
