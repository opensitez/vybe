// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_array_field_reference_cannot_be_replaced_but_elements_can
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder {
    public readonly int[] Data = { 1, 2 };
}
var holder = new Holder();
holder.Data[1] = 9;
__Check((holder.Data[1]).ToString(), "9");
