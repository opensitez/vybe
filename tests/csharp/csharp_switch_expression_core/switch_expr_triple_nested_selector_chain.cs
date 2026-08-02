// vybe-test: csharp/csharp_switch_expression_core/switch_expr_triple_nested_selector_chain
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Depth(int n)=>n switch{1=>2 switch{2=>3 switch{3=>9,_=>0},_=>0},_=>0}; __Check((Depth(1)).ToString(), "0");
