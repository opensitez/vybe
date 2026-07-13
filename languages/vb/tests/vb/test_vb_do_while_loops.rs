use super::helpers::run_vb;

#[test] fn do_while_top_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo While x < 3\nx += 1\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_until_top_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo Until x = 3\nx += 1\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_while_bottom_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo\nx += 1\nLoop While x < 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_until_bottom_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo\nx += 1\nLoop Until x = 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_while_top_no_execute() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 5\nDo While x < 3\nx += 1\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["5"]); }

#[test] fn do_until_top_no_execute() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 5\nDo Until x > 3\nx += 1\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn do_while_bottom_execute_once() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 5\nDo\nx += 1\nLoop While x < 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["6"]); }
#[test] fn do_until_bottom_execute_once() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 5\nDo\nx += 1\nLoop Until x > 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["6"]); }
#[test] fn do_loop_infinite_with_exit() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo\nx += 1\nIf x = 3 Then Exit Do\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_while_continue_do() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0, sum = 0\nDo While x < 4\nx += 1\nIf x = 2 Then Continue Do\nsum += x\nLoop\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"), vec!["8"]); }

#[test] fn while_end_while_basic() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nWhile x < 3\nx += 1\nEnd While\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn while_end_while_no_execute() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 5\nWhile x < 3\nx += 1\nEnd While\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn while_end_while_exit_while() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nWhile True\nx += 1\nIf x = 3 Then Exit While\nEnd While\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn while_end_while_continue_while() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0, sum = 0\nWhile x < 4\nx += 1\nIf x = 2 Then Continue While\nsum += x\nEnd While\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"), vec!["8"]); }
#[test] fn while_wend_legacy() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nWhile x < 3\nx += 1\nEnd While ' Wend is supported in some legacy modes, but End While is standard\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }

#[test] fn do_loop_nested() { assert_eq!(run_vb("Module M\nSub Main()\nDim i = 0, count = 0\nDo While i < 2\nDim j = 0\nDo While j < 3\ncount += 1\nj += 1\nLoop\ni += 1\nLoop\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"), vec!["6"]); }
#[test] fn do_loop_exit_nested() { assert_eq!(run_vb("Module M\nSub Main()\nDim i = 0, count = 0\nDo While i < 3\nDim j = 0\nDo While j < 3\nj += 1\nIf j = 2 Then Exit Do\ncount += 1\nLoop\ni += 1\nLoop\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn while_end_while_nested() { assert_eq!(run_vb("Module M\nSub Main()\nDim i = 0, count = 0\nWhile i < 2\nDim j = 0\nWhile j < 3\ncount += 1\nj += 1\nEnd While\ni += 1\nEnd While\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"), vec!["6"]); }
#[test] fn while_exit_nested() { assert_eq!(run_vb("Module M\nSub Main()\nDim i = 0, count = 0\nWhile i < 3\nDim j = 0\nWhile j < 3\nj += 1\nIf j = 2 Then Exit While\ncount += 1\nEnd While\ni += 1\nEnd While\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_until_bottom_variable_mutation() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo\nx += 1\nIf x = 2 Then x = 5\nLoop Until x >= 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["5"]); }

#[test] fn do_while_true_literal() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo While True\nx += 1\nIf x = 3 Then Exit Do\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_until_false_literal() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 0\nDo Until False\nx += 1\nIf x = 3 Then Exit Do\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn do_loop_condition_function_call() { assert_eq!(run_vb("Module M\nDim calls As Integer = 0\nFunction Check() As Boolean\ncalls += 1\nReturn calls < 3\nEnd Function\nSub Main()\nDo While Check()\nLoop\nConsole.WriteLine(calls)\nEnd Sub\nEnd Module"), vec!["3"]); } // Evaluated each iteration
#[test] fn while_condition_function_call() { assert_eq!(run_vb("Module M\nDim calls As Integer = 0\nFunction Check() As Boolean\ncalls += 1\nReturn calls < 3\nEnd Function\nSub Main()\nWhile Check()\nEnd While\nConsole.WriteLine(calls)\nEnd Sub\nEnd Module"), vec!["3"]); } // Evaluated each iteration
#[test] fn do_loop_boolean_conversion() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim x = 0\nDo While \"True\"\nx += 1\nExit Do\nLoop\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["1"]); } // String "True" coerced to Boolean True
