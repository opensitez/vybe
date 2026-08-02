// vybe-test: csharp/csharp_nameof_expressions/nameof_two_parameters_print_both_names
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Pair(int left,int right){__Check((nameof(left)).ToString(), "left"); __Check((nameof(right)).ToString(), "right");} Pair(1,2);
