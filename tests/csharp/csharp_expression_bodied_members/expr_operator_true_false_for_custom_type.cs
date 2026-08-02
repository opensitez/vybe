// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_true_false_for_custom_type
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

struct Flag { public bool On; public static bool operator true(Flag f) => f.On; public static bool operator false(Flag f) => !f.On; }
Flag f = new Flag { On = true }; if (f) Console.WriteLine("yes"); else Console.WriteLine("no");
