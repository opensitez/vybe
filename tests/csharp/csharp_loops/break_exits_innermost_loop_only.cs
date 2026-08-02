// vybe-test: csharp/csharp_loops/break_exits_innermost_loop_only
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int total = 0;
for(int i=0;i<3;i++) {
    for(int j=0;j<3;j++) {
        if(j==1) break;
        total++;
    }
}
Console.WriteLine(total);
