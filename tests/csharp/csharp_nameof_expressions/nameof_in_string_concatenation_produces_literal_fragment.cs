// vybe-test: csharp/csharp_nameof_expressions/nameof_in_string_concatenation_produces_literal_fragment
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int age=30; __Check(("field="+nameof(age)).ToString(), "field=age");
