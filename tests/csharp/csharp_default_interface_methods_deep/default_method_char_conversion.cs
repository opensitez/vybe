// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_char_conversion
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IChar{char C{get;} string AsString()=>C.ToString();} class Letter:IChar{public char C{get;set;}='Q';} __Check((new Letter().AsString()).ToString(), "Q");
