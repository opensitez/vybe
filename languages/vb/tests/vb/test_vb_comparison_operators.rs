use super::helpers::run_vb;

#[test] fn comp_equal_integers() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 = 5)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_not_equal_integers() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 <> 5)\nEnd Sub\nEnd Module"), vec!["False"]); }
#[test] fn comp_less_than() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 < 10)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_greater_than() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 > 5)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_less_than_or_equal() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 <= 5)\nEnd Sub\nEnd Module"), vec!["True"]); }

#[test] fn comp_greater_than_or_equal() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 >= 6)\nEnd Sub\nEnd Module"), vec!["False"]); }
#[test] fn comp_string_equality() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"A\" = \"A\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_string_not_equal() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"A\" <> \"B\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_is_operator_reference() { assert_eq!(run_vb("Class C\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nDim c2 = c1\nConsole.WriteLine(c1 Is c2)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_is_operator_reference_different() { assert_eq!(run_vb("Class C\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nDim c2 As New C()\nConsole.WriteLine(c1 Is c2)\nEnd Sub\nEnd Module"), vec!["False"]); }

#[test] fn comp_isnot_operator() { assert_eq!(run_vb("Class C\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nDim c2 As New C()\nConsole.WriteLine(c1 IsNot c2)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_is_nothing() { assert_eq!(run_vb("Module M\nSub Main()\nDim obj As Object = Nothing\nConsole.WriteLine(obj Is Nothing)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_isnot_nothing() { assert_eq!(run_vb("Module M\nSub Main()\nDim obj As Object = 1\nConsole.WriteLine(obj IsNot Nothing)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_like_operator_basic() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"Hello\" Like \"H*o\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_like_operator_wildcard_char() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"Bat\" Like \"B?t\")\nEnd Sub\nEnd Module"), vec!["True"]); }

#[test] fn comp_like_operator_digit() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"123\" Like \"1#3\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_like_operator_char_list() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"C\" Like \"[A-Z]\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_like_operator_negated_list() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(\"1\" Like \"[!A-Z]\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_object_equality_value_types() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim obj1 As Object = 10\nDim obj2 As Object = 10\nConsole.WriteLine(obj1 = obj2)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_object_equality_reference_types_throws() { assert_eq!(run_vb("Option Strict Off\nClass C\nEnd Class\nModule M\nSub Main()\nDim c1 As Object = New C()\nDim c2 As Object = New C()\nTry\nConsole.WriteLine(c1 = c2)\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"), vec!["Caught"]); }

#[test] fn comp_type_coercion_string_number() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nConsole.WriteLine(\"10\" = 10)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_boolean_integer() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nConsole.WriteLine(True = -1)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_date_string() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Date = #1/1/2000#\nConsole.WriteLine(d = \"1/1/2000\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_precedence_with_arithmetic() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5 + 5 = 10 And 2 * 2 = 4)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn comp_is_typeof() { assert_eq!(run_vb("Module M\nSub Main()\nDim obj As Object = \"Str\"\nConsole.WriteLine(TypeOf obj Is String)\nEnd Sub\nEnd Module"), vec!["True"]); }
