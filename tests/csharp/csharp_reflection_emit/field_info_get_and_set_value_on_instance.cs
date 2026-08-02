// vybe-test: csharp/csharp_reflection_emit/field_info_get_and_set_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box{public int V;}
var fi=typeof(Box).GetField("V");
var obj=new Box();
fi.SetValue(obj,55);
__Check((fi.GetValue(obj)).ToString(), "55");
