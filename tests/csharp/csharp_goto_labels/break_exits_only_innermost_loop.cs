// vybe-test: csharp/csharp_goto_labels/break_exits_only_innermost_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

int count=0;
for(int i=0;i<3;i++){
    for(int j=0;j<3;j++){
        if(j==1) break;
        count++;
    }
}
Console.WriteLine(count);
