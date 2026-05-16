use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(using_statement_calls_dispose_after_block, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } using (var resource = new Resource()) { Console.WriteLine("inside"); }"#, ["inside", "disposed"]);
csharp_case!(using_statement_disposes_multiple_resources_in_reverse_order, r#"using System; class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { Console.WriteLine(name); } } using (var left = new Resource("left")) using (var right = new Resource("right")) { Console.WriteLine("body"); }"#, ["body", "right", "left"]);
csharp_case!(using_declaration_disposes_at_end_of_scope, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } using var resource = new Resource(); Console.WriteLine("body");"#, ["body", "disposed"]);
csharp_case!(using_statement_runs_dispose_when_exception_is_caught, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } try { using (var resource = new Resource()) { Console.WriteLine("body"); throw new Exception("boom"); } } catch (Exception) { Console.WriteLine("caught"); }"#, ["body", "disposed", "caught"]);
csharp_case!(manual_dispose_method_can_be_called_directly, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } var resource = new Resource(); resource.Dispose();"#, ["disposed"]);
csharp_case!(using_statement_allows_access_to_resource_members_inside_scope, r#"using System; class Resource : IDisposable { public string Name => "file"; public void Dispose() { Console.WriteLine("disposed"); } } using (var resource = new Resource()) { Console.WriteLine(resource.Name); }"#, ["file", "disposed"]);
csharp_case!(nested_using_declarations_dispose_in_reverse_order, r#"using System; class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { Console.WriteLine(name); } } using var first = new Resource("first"); using var second = new Resource("second"); Console.WriteLine("done");"#, ["done", "second", "first"]);
csharp_case!(disposable_can_accumulate_state_before_disposal, r#"using System; class Buffer : IDisposable { int count; public void Add() { count++; } public void Dispose() { Console.WriteLine(count); } } using (var buffer = new Buffer()) { buffer.Add(); buffer.Add(); }"#, ["2"]);
csharp_case!(using_statement_supports_interface_typed_variable, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } using (IDisposable resource = new Resource()) { Console.WriteLine("body"); }"#, ["body", "disposed"]);
csharp_case!(dispose_can_be_invoked_from_helper_method, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } void Close(IDisposable item) { item.Dispose(); } var resource = new Resource(); Close(resource);"#, ["disposed"]);
csharp_case!(try_finally_can_model_manual_cleanup_order, r#"try { Console.WriteLine("body"); } finally { Console.WriteLine("cleanup"); }"#, ["body", "cleanup"]);
csharp_case!(lock_statement_serializes_body_execution, r#"object gate = new object(); lock (gate) { Console.WriteLine("locked"); }"#, ["locked"]);
csharp_case!(using_statement_with_return_still_disposes_resource, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } int Read() { using (var resource = new Resource()) { Console.WriteLine("inside"); return 5; } } Console.WriteLine(Read());"#, ["inside", "disposed", "5"]);
csharp_case!(using_declaration_after_local_function_still_disposes_at_scope_end, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } string Read() { using var resource = new Resource(); return "ok"; } Console.WriteLine(Read());"#, ["disposed", "ok"]);
csharp_case!(multiple_try_finally_blocks_execute_independently, r#"try { Console.WriteLine("one"); } finally { Console.WriteLine("cleanup-one"); } try { Console.WriteLine("two"); } finally { Console.WriteLine("cleanup-two"); }"#, ["one", "cleanup-one", "two", "cleanup-two"]);
csharp_case!(disposable_field_can_be_closed_by_owner_method, r#"using System; class Resource : IDisposable { public void Dispose() { Console.WriteLine("disposed"); } } class Owner { Resource resource = new Resource(); public void Close() { resource.Dispose(); } } new Owner().Close();"#, ["disposed"]);
csharp_case!(lock_statement_can_mutate_shared_local_state, r#"object gate = new object(); int count = 0; lock (gate) { count += 3; } Console.WriteLine(count);"#, ["3"]);
csharp_case!(using_statement_supports_expression_bodied_dispose_member, r#"using System; class Resource : IDisposable { public void Dispose() => Console.WriteLine("disposed"); } using (var resource = new Resource()) { Console.WriteLine("body"); }"#, ["body", "disposed"]);
csharp_case!(finally_block_runs_after_caught_exception, r#"try { throw new System.Exception(); } catch (System.Exception) { Console.WriteLine("caught"); } finally { Console.WriteLine("finally"); }"#, ["caught", "finally"]);
csharp_case!(using_block_can_allocate_and_return_computed_value, r#"using System; class Resource : IDisposable { public int Value => 4; public void Dispose() { Console.WriteLine("disposed"); } } using (var resource = new Resource()) { Console.WriteLine(resource.Value * 2); }"#, ["8", "disposed"]);