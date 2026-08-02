// vybe-test: csharp/csharp_control_flow/foreach_on_list_throws_when_collection_modified_during_iteration
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

var items = new System.Collections.Generic.List<int> { 1, 2 };
string outcome = "ok";
try {
    foreach (var item in items) {
        items.RemoveAt(0);
    }
} catch (System.InvalidOperationException) {
    outcome = "invalid";
}
Console.WriteLine(outcome);
