// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_two_levels_with_relational
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public int N; } class Outer { public Inner I; } object o=new Outer{I=new Inner{N=50}}; __Check((o is Outer{I:{N:>40}}).ToString(), "True");
