// vybe-test: csharp/csharp_exception_types/argument_exception_carries_param_name
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { throw new System.ArgumentException("bad","myParam"); }
catch(System.ArgumentException e) { __Check((e.ParamName).ToString(), "myParam"); }
