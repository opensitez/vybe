// vybe-test: csharp/csharp_pattern_property/switch_statement_property_pattern_case_with_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node { public int Id; } object o=new Node{Id=12}; string tag=""; switch(o){case Node{Id:var id}:tag=id.ToString();break;default:tag="0";break;} __Check((tag).ToString(), "12");
