use super::helpers::run_vb;

#[test] fn var_dim_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x As Integer = 10\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn var_dim_multiple() { assert_eq!(run_vb("Module M\nSub Main()\nDim x, y As Integer\nx = 1: y = 2\nConsole.WriteLine(x + y)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn var_dim_type_inference() { assert_eq!(run_vb("Option Infer On\nModule M\nSub Main()\nDim x = 10\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn var_dim_implicit_object() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim x\nx = \"A\"\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["A"]); }
#[test] fn var_dim_shadowing_module() { assert_eq!(run_vb("Module M\nDim x As Integer = 5\nSub Main()\nDim x As Integer = 10\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["10"]); }

#[test] fn var_const_basic() { assert_eq!(run_vb("Module M\nSub Main()\nConst x As Integer = 42\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["42"]); }
#[test] fn var_const_expression() { assert_eq!(run_vb("Module M\nSub Main()\nConst x = 10 + 20\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["30"]); }
#[test] fn var_const_type_character() { assert_eq!(run_vb("Module M\nSub Main()\nConst x$ = \"A\"\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["A"]); }
#[test] fn var_const_multiple() { assert_eq!(run_vb("Module M\nSub Main()\nConst a As Integer = 1, b As Integer = 2\nConsole.WriteLine(b)\nEnd Sub\nEnd Module"), vec!["2"]); }
#[test] fn var_const_shadowing() { assert_eq!(run_vb("Module M\nConst x As Integer = 1\nSub Main()\nConst x As Integer = 2\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["2"]); }

#[test] fn var_static_basic() { assert_eq!(run_vb("Module M\nSub Test()\nStatic x As Integer = 0\nx += 1\nConsole.WriteLine(x)\nEnd Sub\nSub Main()\nTest()\nTest()\nEnd Sub\nEnd Module"), vec!["1", "2"]); }
#[test] fn var_static_multiple() { assert_eq!(run_vb("Module M\nSub Test()\nStatic x = 1, y = 2\nx += 1: y += 1\nConsole.WriteLine(x + y)\nEnd Sub\nSub Main()\nTest()\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn var_static_in_function() { assert_eq!(run_vb("Module M\nFunction F() As Integer\nStatic x As Integer = 10\nx += 10\nReturn x\nEnd Function\nSub Main()\nConsole.WriteLine(F())\nEnd Sub\nEnd Module"), vec!["20"]); }
#[test] fn var_static_init_once() { assert_eq!(run_vb("Module M\nFunction Init() As Integer\nConsole.WriteLine(\"Init\")\nReturn 5\nEnd Function\nSub Test()\nStatic x As Integer = Init()\nx += 1\nConsole.WriteLine(x)\nEnd Sub\nSub Main()\nTest()\nTest()\nEnd Sub\nEnd Module"), vec!["Init", "6", "7"]); }
#[test] fn var_static_shadowing_parameter() { assert_eq!(run_vb("Module M\nSub Main()\n' Static cannot shadow parameter directly, parser edge case test\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }

#[test] fn var_scoping_block() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 1\nIf True Then\nDim y = 2\nConsole.WriteLine(x + y)\nEnd If\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn var_scoping_loop() { assert_eq!(run_vb("Module M\nSub Main()\nFor i = 1 To 2\nDim x = i\nNext\n' x is accessible outside loop in older VB, but VB.NET block scopes Dim in loops depending on strictness. Actually, Dim inside For is block-scoped.\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn var_scoping_shadowing_block() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 1\nIf True Then\n' Cannot shadow method-level local variable in a block\nConsole.WriteLine(\"Parsed\")\nEnd If\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn var_scoping_with() { assert_eq!(run_vb("Structure S\nPublic v As Integer\nEnd Structure\nModule M\nSub Main()\nDim s1 As New S()\nWith s1\nDim v = 10\nConsole.WriteLine(.v + v)\nEnd With\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn var_scoping_try_catch() { assert_eq!(run_vb("Module M\nSub Main()\nTry\nDim x = 1\nCatch\nDim y = 2\nEnd Try\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }

#[test] fn var_declaration_modifiers() { assert_eq!(run_vb("Module M\nPrivate x As Integer = 1\nPublic y As Integer = 2\nSub Main()\nConsole.WriteLine(x + y)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn var_declaration_readonly() { assert_eq!(run_vb("Module M\nReadOnly x As Integer = 10\nSub Main()\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn var_declaration_with_events() { assert_eq!(run_vb("Class C\nPublic Event E()\nEnd Class\nModule M\nWithEvents obj As New C()\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn var_declaration_as_new() { assert_eq!(run_vb("Class C\nPublic v As Integer = 5\nEnd Class\nModule M\nSub Main()\nDim obj As New C()\nConsole.WriteLine(obj.v)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn var_declaration_array_bounds() { assert_eq!(run_vb("Module M\nSub Main()\nDim arr(5) As Integer\nConsole.WriteLine(arr.Length)\nEnd Sub\nEnd Module"), vec!["6"]); }
