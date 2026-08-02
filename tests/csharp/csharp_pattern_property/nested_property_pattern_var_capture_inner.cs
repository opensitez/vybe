// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_var_capture_inner
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public int N; } class Outer { public Inner Child; } object o=new Outer{Child=new Inner{N=9}}; if(o is Outer{Child:{N:var n}}) __Check((n).ToString(), "9");
