// vybe-test: csharp/csharp_pattern_property/switch_statement_property_pattern_case
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node { public int Id; } object o=new Node{Id=5}; string tag=""; switch(o){case Node{Id:5}:tag="match";break;default:tag="miss";break;} __Check((tag).ToString(), "match");
