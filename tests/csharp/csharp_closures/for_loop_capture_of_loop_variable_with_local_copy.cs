// vybe-test: csharp/csharp_closures/for_loop_capture_of_loop_variable_with_local_copy
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

var actions = new System.Collections.Generic.List<System.Func<int>>();
for(int i=0; i<3; i++) {
    var copy = i;
    actions.Add(() => copy);
}
foreach(var a in actions) Console.WriteLine(a());
