// vybe-test: csharp/csharp_diagnostics_process_start_info/process_start_info_case_15

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var psi = new System.Diagnostics.ProcessStartInfo("dotnet", "--version");
psi.EnvironmentVariables["TEST_VAR_15"] = "Val_15";
__P(psi.FileName);
__P(psi.Arguments);
__P(psi.EnvironmentVariables["TEST_VAR_15"]);
__Check("dotnet\n--version\nVal_15");
