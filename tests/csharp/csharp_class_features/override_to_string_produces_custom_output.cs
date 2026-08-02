// vybe-test: csharp/csharp_class_features/override_to_string_produces_custom_output
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Color{int R,G,B;public Color(int r,int g,int b){R=r;G=g;B=b;}
public override string ToString()=>$"rgb({R},{G},{B})";}
__Check((new Color(255,0,128)).ToString(), "rgb(255,0,128)");
