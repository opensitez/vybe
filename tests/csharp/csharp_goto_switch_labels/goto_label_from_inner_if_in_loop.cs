// vybe-test: csharp/csharp_goto_switch_labels/goto_label_from_inner_if_in_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int sum = 0;
for (int i = 0; i < 5; i++) {
    if (i == 3) goto done;
    sum += i;
}
done:
Console.WriteLine(sum);
