// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_matches_inner_field
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public int N; } class Outer { public Inner Child; } object o=new Outer{Child=new Inner{N=7}}; if(o is Outer{Child:{N:7}}) __Check(("ok").ToString(), "ok");
