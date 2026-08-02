// vybe-test: csharp/csharp_goto_labels/continue_skips_rest_of_current_iteration
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

int sum=0;
for(int i=1;i<=10;i++){
    if(i%2==0) continue;
    sum+=i;
}
Console.WriteLine(sum);
