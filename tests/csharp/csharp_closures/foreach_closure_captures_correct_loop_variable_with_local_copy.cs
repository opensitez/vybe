// vybe-test: csharp/csharp_closures/foreach_closure_captures_correct_loop_variable_with_local_copy
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

var actions = new System.Collections.Generic.List<System.Func<int>>();
foreach(var v in new[]{10,20,30}) {
    var copy = v;
    actions.Add(() => copy);
}
foreach(var a in actions) Console.WriteLine(a());
