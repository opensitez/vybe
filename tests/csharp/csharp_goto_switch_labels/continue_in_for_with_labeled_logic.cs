// vybe-test: csharp/csharp_goto_switch_labels/continue_in_for_with_labeled_logic
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string chars = "";
for (int i = 0; i < 4; i++) {
    if (i == 2) continue;
    chars += i.ToString();
}
Console.WriteLine(chars);
