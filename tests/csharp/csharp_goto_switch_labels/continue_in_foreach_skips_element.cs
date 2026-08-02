// vybe-test: csharp/csharp_goto_switch_labels/continue_in_foreach_skips_element
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int sum = 0;
foreach (var x in new int[] { 1, 2, 3, 4 }) {
    if (x == 2) continue;
    sum += x;
}
Console.WriteLine(sum);
