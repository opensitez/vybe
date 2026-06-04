use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    local_function_returns_sum_of_two_numbers,
    r#"int Add(int left, int right) { return left + right; } Console.WriteLine(Add(3, 4));"#,
    ["7"]
);
csharp_case!(
    local_function_captures_enclosing_variable,
    r#"int factor = 3; int Scale(int value) { return value * factor; } Console.WriteLine(Scale(5));"#,
    ["15"]
);
csharp_case!(
    local_function_can_be_recursive,
    r#"int Fib(int n) { return n <= 1 ? n : Fib(n - 1) + Fib(n - 2); } Console.WriteLine(Fib(6));"#,
    ["8"]
);
csharp_case!(
    local_function_with_expression_body_returns_value,
    r#"int Double(int value) => value * 2; Console.WriteLine(Double(9));"#,
    ["18"]
);
csharp_case!(
    local_function_can_return_tuple_result,
    r#"(int, int) Pair() { return (2, 5); } var result = Pair(); Console.WriteLine(result.Item1 + result.Item2);"#,
    ["7"]
);
csharp_case!(
    local_function_can_use_default_parameter_value,
    r#"int Add(int left, int right = 10) { return left + right; } Console.WriteLine(Add(5));"#,
    ["15"]
);
csharp_case!(
    local_function_can_write_to_outer_variable,
    r#"int total = 0; void Add(int value) { total += value; } Add(4); Add(6); Console.WriteLine(total);"#,
    ["10"]
);
csharp_case!(
    local_function_inside_loop_can_access_iteration_value,
    r#"foreach (var item in new[] { 1, 2, 3 }) { int Square() { return item * item; } Console.WriteLine(Square()); }"#,
    ["1", "4", "9"]
);
csharp_case!(
    local_function_can_be_called_before_its_declaration,
    r#"Console.WriteLine(Read()); string Read() { return "ok"; }"#,
    ["ok"]
);
csharp_case!(
    static_local_function_does_not_capture_outer_state,
    r#"static int Triple(int value) { return value * 3; } Console.WriteLine(Triple(4));"#,
    ["12"]
);
csharp_case!(
    generic_local_function_returns_typed_argument,
    r#"T Echo<T>(T value) { return value; } Console.WriteLine(Echo("generic"));"#,
    ["generic"]
);
csharp_case!(
    local_function_can_have_out_parameter,
    r#"void Split(int value, out int left, out int right) { left = value / 2; right = value - left; } Split(9, out var left, out var right); Console.WriteLine(left); Console.WriteLine(right);"#,
    ["4", "5"]
);
csharp_case!(
    local_function_can_have_ref_parameter,
    r#"void Increment(ref int value) { value++; } int count = 7; Increment(ref count); Console.WriteLine(count);"#,
    ["8"]
);
csharp_case!(
    partial_method_implemented_in_second_part_can_be_called,
    r#"partial class Worker { partial void OnRun(); public void Run() { OnRun(); } } partial class Worker { partial void OnRun() { System.Console.WriteLine("ran"); } } new Worker().Run();"#,
    ["ran"]
);
csharp_case!(
    partial_method_can_receive_argument_from_first_part,
    r#"partial class Worker { partial void OnRun(int value); public void Run() { OnRun(5); } } partial class Worker { partial void OnRun(int value) { System.Console.WriteLine(value * 2); } } new Worker().Run();"#,
    ["10"]
);
csharp_case!(
    partial_method_can_access_shared_private_field,
    r#"partial class Worker { int count = 3; partial void OnRun(); public void Run() { OnRun(); } } partial class Worker { partial void OnRun() { System.Console.WriteLine(count); } } new Worker().Run();"#,
    ["3"]
);
csharp_case!(
    partial_method_can_be_triggered_from_constructor,
    r#"partial class Worker { partial void OnCreated(); public Worker() { OnCreated(); } } partial class Worker { partial void OnCreated() { System.Console.WriteLine("created"); } } new Worker();"#,
    ["created"]
);
csharp_case!(
    partial_method_can_be_invoked_multiple_times,
    r#"partial class Worker { partial void OnRun(); public void RunTwice() { OnRun(); OnRun(); } } partial class Worker { partial void OnRun() { System.Console.WriteLine("tick"); } } new Worker().RunTwice();"#,
    ["tick", "tick"]
);
csharp_case!(
    local_function_can_return_lambda_result,
    r#"int Compute() { System.Func<int> read = () => 9; return read() + 1; } Console.WriteLine(Compute());"#,
    ["10"]
);
csharp_case!(
    local_function_can_pattern_match_on_argument_type,
    r#"string Describe(object value) { return value is int number ? (number * 2).ToString() : "other"; } Console.WriteLine(Describe(6));"#,
    ["12"]
);
