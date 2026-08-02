// vybe-test: csharp/csharp_exception_types/argument_null_exception_message_contains_param_name
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { throw new System.ArgumentNullException("value"); }
catch(System.ArgumentNullException e) { __Check((e.ParamName).ToString(), "value"); }
