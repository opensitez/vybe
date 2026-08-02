// vybe-test: csharp/csharp_pattern_property/is_property_pattern_var_capture_reads_field
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Value; } object o=new Box{Value=25}; if(o is Box{Value:var v}) __Check((v).ToString(), "25");
