// vybe-test: csharp/csharp_loops/nested_foreach_produces_cartesian_pair_count
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int count=0;
foreach(var a in new[]{1,2})
    foreach(var b in new[]{1,2,3})
        count++;
Console.WriteLine(count);
