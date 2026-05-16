use super::helpers::{load_vb_profile, run_vb, run_vb_vm};
use vybe_bytecode::{Value, value::ObjectKind};

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_full_spec!(delegate_spec_single_line_function_lambda_doubles_input, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer) = Function(x) x * 2
        Console.WriteLine(fn(5))
    End Sub
End Module"#, ["10"]);
vb_full_spec!(delegate_spec_multiline_function_lambda_returns_branch_value, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, String) = Function(x)
            If x > 0 Then
                Return "positive"
            End If
            Return "nonpositive"
        End Function
        Console.WriteLine(fn(2))
    End Sub
End Module"#, ["positive"]);
vb_full_spec!(delegate_spec_sub_lambda_prints_message, r#"Module M
    Sub Main()
        Dim action As Action = Sub() Console.WriteLine("hi")
        action()
    End Sub
End Module"#, ["hi"]);
vb_full_spec!(delegate_spec_lambda_captures_local_variable, r#"Module M
    Sub Main()
        Dim prefix As String = "hello"
        Dim fn As Func(Of String) = Function() prefix
        Console.WriteLine(fn())
    End Sub
End Module"#, ["hello"]);
vb_full_spec!(delegate_spec_lambda_updates_captured_variable, r#"Module M
    Sub Main()
        Dim total As Integer = 0
        Dim action As Action = Sub() total += 1
        action()
        action()
        Console.WriteLine(total)
    End Sub
End Module"#, ["2"]);
vb_full_spec!(delegate_spec_lambda_can_be_passed_as_argument, r#"Module M
    Function Apply(fn As Func(Of Integer, Integer), value As Integer) As Integer
        Return fn(value)
    End Function
    Sub Main()
        Console.WriteLine(Apply(Function(x) x + 3, 4))
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_lambda_can_be_returned_from_function, r#"Module M
    Function MakeAdder(seed As Integer) As Func(Of Integer, Integer)
        Return Function(x) x + seed
    End Function
    Sub Main()
        Dim addFive = MakeAdder(5)
        Console.WriteLine(addFive(7))
    End Sub
End Module"#, ["12"]);
vb_full_spec!(delegate_spec_lambda_can_return_another_lambda, r#"Module M
    Function MakeFactory() As Func(Of Func(Of Integer, Integer))
        Return Function() Function(x) x * 3
    End Function
    Sub Main()
        Console.WriteLine(MakeFactory()()(4))
    End Sub
End Module"#, ["12"]);
vb_full_spec!(delegate_spec_addressof_module_function_matches_func_delegate, r#"Module M
    Function Square(x As Integer) As Integer
        Return x * x
    End Function
    Sub Main()
        Dim fn As Func(Of Integer, Integer) = AddressOf Square
        Console.WriteLine(fn(6))
    End Sub
End Module"#, ["36"]);
vb_full_spec!(delegate_spec_addressof_instance_method_matches_func_delegate, r#"Class Counter
    Public Function DoubleValue(x As Integer) As Integer
        Return x * 2
    End Function
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        Dim fn As Func(Of Integer, Integer) = AddressOf c.DoubleValue
        Console.WriteLine(fn(8))
    End Sub
End Module"#, ["16"]);
vb_full_spec!(delegate_spec_addressof_shared_method_matches_func_delegate, r#"Class MathBox
    Public Shared Function Triple(x As Integer) As Integer
        Return x * 3
    End Function
End Class
Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer) = AddressOf MathBox.Triple
        Console.WriteLine(fn(4))
    End Sub
End Module"#, ["12"]);
vb_full_spec!(delegate_spec_action_delegate_can_point_to_sub, r#"Module M
    Sub Speak()
        Console.WriteLine("speak")
    End Sub
    Sub Main()
        Dim action As Action = AddressOf Speak
        action()
    End Sub
End Module"#, ["speak"]);
vb_full_spec!(delegate_spec_predicate_delegate_returns_boolean_result, r#"Module M
    Sub Main()
        Dim fn As Predicate(Of Integer) = Function(x) x > 10
        Console.WriteLine(fn(12))
    End Sub
End Module"#, ["true"]);
vb_full_spec!(delegate_spec_func_with_two_parameters_returns_sum, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer, Integer) = Function(a, b) a + b
        Console.WriteLine(fn(3, 4))
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_func_with_three_parameters_returns_total, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer, Integer, Integer) = Function(a, b, c) a + b + c
        Console.WriteLine(fn(1, 2, 3))
    End Sub
End Module"#, ["6"]);
vb_full_spec!(delegate_spec_lambda_can_close_over_parameter_from_factory, r#"Module M
    Function Make(prefix As String) As Func(Of String)
        Return Function() prefix & "!"
    End Function
    Sub Main()
        Console.WriteLine(Make("vb")())
    End Sub
End Module"#, ["vb!"]);
vb_full_spec!(delegate_spec_lambda_inside_loop_can_use_iteration_value, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer)
        For i As Integer = 1 To 3
            fn = Function(x) x + i
        Next
        Console.WriteLine(fn(4))
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_lambda_can_capture_reference_type_and_mutate_field, r#"Class Box
    Public Value As Integer
End Class
Module M
    Sub Main()
        Dim box As New Box()
        Dim action As Action = Sub() box.Value += 2
        action()
        action()
        Console.WriteLine(box.Value)
    End Sub
End Module"#, ["4"]);
vb_full_spec!(delegate_spec_lambda_can_be_stored_in_list_of_actions, r#"Module M
    Sub Main()
        Dim actions As New List(Of Action)
        actions.Add(Sub() Console.WriteLine("a"))
        actions.Add(Sub() Console.WriteLine("b"))
        actions(0)()
        actions(1)()
    End Sub
End Module"#, ["a", "b"]);
vb_full_spec!(delegate_spec_lambda_array_can_store_multiple_functions, r#"Module M
    Sub Main()
        Dim funcs() As Func(Of Integer, Integer) = {Function(x) x + 1, Function(x) x + 2}
        Console.WriteLine(funcs(1)(5))
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_delegate_invocation_list_runs_multicast_subs, r#"Module M
    Sub Main()
        Dim action As Action = Sub() Console.WriteLine("a")
        action = CType([Delegate].Combine(action, CType(Sub() Console.WriteLine("b"), Action)), Action)
        action()
    End Sub
End Module"#, ["a", "b"]);

#[test]
fn delegate_spec_delegate_combine_stores_multicast_array_value() {
    let (vm, _) = run_vb_vm(r#"Module M
    Dim action As Action
    Sub Main()
        action = Sub() Console.WriteLine("a")
        action = CType([Delegate].Combine(action, CType(Sub() Console.WriteLine("b"), Action)), Action)
    End Sub
End Module"#);

    let action = vm.globals.get("action").expect("expected module-level action global");
    let Value::Object(action_obj) = action else {
        panic!("expected action global to be an object, got {action:?}");
    };
    let action_guard = action_obj.lock().unwrap();
    let ObjectKind::Array(handlers) = &action_guard.kind else {
        panic!("expected combined action to be an array invocation list, got {:?}", action_guard.kind);
    };
    assert_eq!(handlers.len(), 2, "expected exactly two handlers in the invocation list");
    for handler in handlers {
        let Value::Object(handler_obj) = handler else {
            panic!("expected multicast handler to be an object, got {handler:?}");
        };
        let handler_guard = handler_obj.lock().unwrap();
        assert!(matches!(handler_guard.kind, ObjectKind::Function(_)), "expected multicast handler to be a function, got {:?}", handler_guard.kind);
    }
}

#[test]
fn delegate_spec_delegate_combine_local_without_invocation_does_not_overflow() {
    let output = run_vb(r#"Module M
    Sub Main()
        Dim action As Action = Sub() Console.WriteLine("a")
        action = CType([Delegate].Combine(action, CType(Sub() Console.WriteLine("b"), Action)), Action)
    End Sub
End Module"#);

    assert!(output.is_empty(), "unexpected output while combining delegates: {output:?}");
}

#[test]
fn delegate_spec_delegate_combine_compiles_without_overflow() {
    let module = vybe_compiler::languages::vb::parse(r#"Module M
    Sub Main()
        Dim action As Action = Sub() Console.WriteLine("a")
        action = CType([Delegate].Combine(action, CType(Sub() Console.WriteLine("b"), Action)), Action)
    End Sub
End Module"#).expect("VB parse failed");

    let profile = load_vb_profile();
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");
}

#[test]
fn delegate_spec_ctype_sub_lambda_compiles_without_overflow() {
    let module = vybe_compiler::languages::vb::parse(r#"Module M
    Sub Main()
        Dim action As Action = CType(Sub() Console.WriteLine("b"), Action)
    End Sub
End Module"#).expect("VB parse failed");

    let profile = load_vb_profile();
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");
}

#[test]
fn delegate_spec_delegate_combine_with_casted_sub_lambda_compiles_without_overflow() {
    let module = vybe_compiler::languages::vb::parse(r#"Module M
    Sub Main()
        Dim action As Action = Sub() Console.WriteLine("a")
        Dim combined = [Delegate].Combine(action, CType(Sub() Console.WriteLine("b"), Action))
    End Sub
End Module"#).expect("VB parse failed");

    let profile = load_vb_profile();
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");
}
vb_full_spec!(delegate_spec_addhandler_with_lambda_receives_event, r#"Class Clock
    Public Event Tick(value As Integer)
    Public Sub RaiseTick(value As Integer)
        RaiseEvent Tick(value)
    End Sub
End Class
Module M
    Sub Main()
        Dim clock As New Clock()
        AddHandler clock.Tick, Sub(value As Integer) Console.WriteLine(value)
        clock.RaiseTick(9)
    End Sub
End Module"#, ["9"]);
vb_full_spec!(delegate_spec_removehandler_with_lambda_variable_detaches_listener, r#"Class Clock
    Public Event Tick()
    Public Sub RaiseTick()
        RaiseEvent Tick()
    End Sub
End Class
Module M
    Sub Main()
        Dim clock As New Clock()
        Dim handler As Action = Sub() Console.WriteLine("tick")
        AddHandler clock.Tick, handler
        RemoveHandler clock.Tick, handler
        clock.RaiseTick()
        Console.WriteLine("done")
    End Sub
End Module"#, ["done"]);
vb_full_spec!(delegate_spec_custom_delegate_type_can_be_invoked, r#"Delegate Function Combiner(a As Integer, b As Integer) As Integer
Module M
    Sub Main()
        Dim fn As Combiner = Function(a, b) a + b
        Console.WriteLine(fn(2, 5))
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_delegate_parameter_can_be_called_twice, r#"Module M
    Function ApplyTwice(fn As Func(Of Integer, Integer), value As Integer) As Integer
        Return fn(fn(value))
    End Function
    Sub Main()
        Console.WriteLine(ApplyTwice(Function(x) x + 1, 3))
    End Sub
End Module"#, ["5"]);
vb_full_spec!(delegate_spec_delegate_return_value_can_flow_into_expression, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer) = Function() 7
        Console.WriteLine(fn() + 3)
    End Sub
End Module"#, ["10"]);
vb_full_spec!(delegate_spec_lambda_can_call_private_helper_function, r#"Module M
    Private Function DoubleValue(x As Integer) As Integer
        Return x * 2
    End Function
    Sub Main()
        Dim fn As Func(Of Integer, Integer) = Function(x) DoubleValue(x)
        Console.WriteLine(fn(9))
    End Sub
End Module"#, ["18"]);
vb_full_spec!(delegate_spec_lambda_can_reference_me_member_inside_instance_method, r#"Class Counter
    Public Value As Integer = 4
    Public Function Build() As Func(Of Integer)
        Return Function() Me.Value
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Counter()).Build()())
    End Sub
End Module"#, ["4"]);
vb_full_spec!(delegate_spec_lambda_can_infer_parameter_type_from_delegate, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer) = Function(x) x + 4
        Console.WriteLine(fn(6))
    End Sub
End Module"#, ["10"]);
vb_full_spec!(delegate_spec_multiline_sub_lambda_can_mutate_two_values, r#"Module M
    Sub Main()
        Dim left As Integer = 1
        Dim right As Integer = 2
        Dim action As Action = Sub()
            left += 2
            right += 3
        End Sub
        action()
        Console.WriteLine(left + right)
    End Sub
End Module"#, ["8"]);
vb_full_spec!(delegate_spec_lambda_can_return_boolean_comparison, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, Boolean) = Function(x) x >= 10
        Console.WriteLine(fn(10))
    End Sub
End Module"#, ["true"]);
vb_full_spec!(delegate_spec_lambda_can_call_method_with_optional_parameter, r#"Module M
    Function Show(Optional value As Integer = 9) As Integer
        Return value
    End Function
    Sub Main()
        Dim fn As Func(Of Integer) = Function() Show()
        Console.WriteLine(fn())
    End Sub
End Module"#, ["9"]);
vb_full_spec!(delegate_spec_lambda_can_be_selected_with_if_operator, r#"Module M
    Sub Main()
        Dim yesFn As Func(Of Integer) = Function() 1
        Dim noFn As Func(Of Integer) = Function() 2
        Dim chosen = If(True, yesFn, noFn)
        Console.WriteLine(chosen())
    End Sub
End Module"#, ["1"]);
vb_full_spec!(delegate_spec_lambda_can_return_string_concatenation, r#"Module M
    Sub Main()
        Dim fn As Func(Of String, String, String) = Function(a, b) a & b
        Console.WriteLine(fn("v", "b"))
    End Sub
End Module"#, ["vb"]);
vb_full_spec!(delegate_spec_lambda_can_bind_to_event_style_signature, r#"Module M
    Sub Main()
        Dim handler As Action(Of Object, EventArgs) = Sub(sender As Object, e As EventArgs)
            Console.WriteLine("handled")
        End Sub
        handler(Nothing, Nothing)
    End Sub
End Module"#, ["handled"]);
vb_full_spec!(delegate_spec_nested_lambda_can_capture_outer_lambda_variable, r#"Module M
    Sub Main()
        Dim outer As Func(Of Integer, Func(Of Integer)) = Function(seed)
            Return Function() seed + 1
        End Function
        Console.WriteLine(outer(8)())
    End Sub
End Module"#, ["9"]);
vb_full_spec!(delegate_spec_function_delegate_can_be_passed_through_function, r#"Module M
    Function PassThrough(fn As Func(Of Integer, Integer)) As Func(Of Integer, Integer)
        Return fn
    End Function
    Sub Main()
        Dim fn = PassThrough(Function(x) x * 4)
        Console.WriteLine(fn(3))
    End Sub
End Module"#, ["12"]);
vb_full_spec!(delegate_spec_action_delegate_can_be_returned_from_factory, r#"Module M
    Function MakeAction(text As String) As Action
        Return Sub() Console.WriteLine(text)
    End Function
    Sub Main()
        Dim action = MakeAction("factory")
        action()
    End Sub
End Module"#, ["factory"]);
vb_full_spec!(delegate_spec_lambda_can_capture_array_and_index_into_it, r#"Module M
    Sub Main()
        Dim items() As Integer = {3, 5, 7}
        Dim fn As Func(Of Integer) = Function() items(1)
        Console.WriteLine(fn())
    End Sub
End Module"#, ["5"]);
vb_full_spec!(delegate_spec_lambda_can_capture_dictionary_and_read_key, r#"Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)
        map.Add("x", 9)
        Dim fn As Func(Of Integer) = Function() map("x")
        Console.WriteLine(fn())
    End Sub
End Module"#, ["9"]);
vb_full_spec!(delegate_spec_lambda_can_short_circuit_with_inline_if, r#"Module M
    Sub Main()
        Dim fn As Func(Of Integer, String) = Function(x) If(x > 0, "pos", "neg")
        Console.WriteLine(fn(1))
    End Sub
End Module"#, ["pos"]);
vb_full_spec!(delegate_spec_delegate_can_target_constructor_via_lambda, r#"Class Box
    Public Value As Integer
    Public Sub New(value As Integer)
        Me.Value = value
    End Sub
End Class
Module M
    Sub Main()
        Dim fn As Func(Of Integer, Box) = Function(x) New Box(x)
        Console.WriteLine(fn(6).Value)
    End Sub
End Module"#, ["6"]);
vb_full_spec!(delegate_spec_lambda_can_capture_nullable_value, r#"Module M
    Sub Main()
        Dim value As Integer? = 7
        Dim fn As Func(Of Integer?) = Function() value
        Console.WriteLine(fn())
    End Sub
End Module"#, ["7"]);
vb_full_spec!(delegate_spec_lambda_can_return_object_reference, r#"Class Box
    Public Value As String = "box"
End Class
Module M
    Sub Main()
        Dim fn As Func(Of Box) = Function() New Box()
        Console.WriteLine(fn().Value)
    End Sub
End Module"#, ["box"]);
vb_full_spec!(delegate_spec_delegate_can_chain_two_transform_functions, r#"Module M
    Function Compose(firstFn As Func(Of Integer, Integer), secondFn As Func(Of Integer, Integer), value As Integer) As Integer
        Return secondFn(firstFn(value))
    End Function
    Sub Main()
        Console.WriteLine(Compose(Function(x) x + 1, Function(x) x * 2, 4))
    End Sub
End Module"#, ["10"]);
vb_full_spec!(delegate_spec_lambda_can_capture_loop_built_string, r#"Module M
    Sub Main()
        Dim text As String = ""
        For i As Integer = 1 To 3
            text &= i
        Next
        Dim fn As Func(Of String) = Function() text
        Console.WriteLine(fn())
    End Sub
End Module"#, ["123"]);
vb_full_spec!(delegate_spec_delegate_can_invoke_instance_method_after_state_change, r#"Class Counter
    Public Value As Integer
    Public Function Read() As Integer
        Return Value
    End Function
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        Dim fn As Func(Of Integer) = AddressOf c.Read
        c.Value = 11
        Console.WriteLine(fn())
    End Sub
End Module"#, ["11"]);
vb_full_spec!(delegate_spec_function_lambda_can_return_array_value, r#"Module M
    Sub Main()
        Dim items() As Integer = {2, 4, 6}
        Dim fn As Func(Of Integer) = Function() items(2)
        Console.WriteLine(fn())
    End Sub
End Module"#, ["6"]);
vb_full_spec!(delegate_spec_action_lambda_can_append_to_list, r#"Module M
    Sub Main()
        Dim items As New List(Of Integer)
        Dim action As Action = Sub() items.Add(5)
        action()
        Console.WriteLine(items.Count)
    End Sub
End Module"#, ["1"]);
vb_full_spec!(delegate_spec_lambda_can_project_object_field, r#"Class Box
    Public Value As Integer = 12
End Class
Module M
    Sub Main()
        Dim box As New Box()
        Dim fn As Func(Of Integer) = Function() box.Value
        Console.WriteLine(fn())
    End Sub
End Module"#, ["12"]);
