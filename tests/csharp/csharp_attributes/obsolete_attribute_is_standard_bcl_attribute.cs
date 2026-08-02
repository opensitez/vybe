// vybe-test: csharp/csharp_attributes/obsolete_attribute_is_standard_bcl_attribute
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Old{
    [System.Obsolete("use NewMethod")]
    public void OldMethod(){}
}
var mi=typeof(Old).GetMethod("OldMethod");
bool hasObs=mi.GetCustomAttributes(typeof(System.ObsoleteAttribute),false).Length>0;
__Check((hasObs).ToString(), "True");
