// vybe-test: csharp/csharp_environment_variables/deleting_environment_variable_makes_it_null
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY","x");
System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY",null);
__P((System.Environment.GetEnvironmentVariable("VYBE_DEL_KEY")==null).ToString());
__Check("True");
