// vybe-test: csharp/csharp_records_advanced/record_with_expression_keeps_untouched_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

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

record Item(string Name, int Count); var item = new Item("pen", 2); var changed = item with { Count = 5 }; __P((changed.Name).ToString());
__Check("pen");
