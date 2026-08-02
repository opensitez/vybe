// vybe-test: csharp/csharp_goto_switch_labels/foreach_break_exits_after_first_match
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string seen = "";
foreach (var ch in "abc") {
    seen += ch;
    if (ch == 'b') break;
}
Console.WriteLine(seen);
