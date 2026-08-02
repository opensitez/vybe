// vybe-test: csharp/classes/delegate_declaration_parses
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

delegate int MathOp(int a, int b);
        __Check(("parsed").ToString(), "parsed");
