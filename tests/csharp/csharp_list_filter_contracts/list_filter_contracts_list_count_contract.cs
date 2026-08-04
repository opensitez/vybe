// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

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

// list_filter_contracts
var values = new System.Collections.Generic.List<int> { 31, 32, 31 }; __P((values.Count == 3).ToString());
__Check("True");
