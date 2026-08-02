// vybe-test: csharp/csharp_yield_advanced/yield_in_try_finally_disposes_after_iteration
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

bool cleaned=false;
System.Collections.Generic.IEnumerable<int> Gen(){
    try{ yield return 1; yield return 2; }
    finally{ cleaned=true; }
}
foreach(var _ in Gen()){}
Console.WriteLine(cleaned);
