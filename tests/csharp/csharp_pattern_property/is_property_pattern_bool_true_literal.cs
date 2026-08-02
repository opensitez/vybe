// vybe-test: csharp/csharp_pattern_property/is_property_pattern_bool_true_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Flag { public bool On; } object o=new Flag{On=true}; __Check((o is Flag{On:true}).ToString(), "True");
