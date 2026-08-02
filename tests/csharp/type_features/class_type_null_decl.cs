// vybe-test: csharp/type_features/class_type_null_decl
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bar { public int value; public Bar(int v) { this.value = v; } }
        Bar b = null;
        __Check((b?.value ?? "none").ToString(), "none");
