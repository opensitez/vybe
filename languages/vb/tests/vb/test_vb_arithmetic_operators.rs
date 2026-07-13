use super::helpers::run_vb;

#[test] fn arith_addition() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 + 20)\nEnd Sub\nEnd Module"), vec!["30"]); }
#[test] fn arith_subtraction() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(20 - 5)\nEnd Sub\nEnd Module"), vec!["15"]); }
#[test] fn arith_multiplication() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 * 3)\nEnd Sub\nEnd Module"), vec!["30"]); }
#[test] fn arith_division_normal() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 / 4)\nEnd Sub\nEnd Module"), vec!["2.5"]); }
#[test] fn arith_division_integer() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 \\ 3)\nEnd Sub\nEnd Module"), vec!["3"]); } // Integer division

#[test] fn arith_modulo() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(10 Mod 3)\nEnd Sub\nEnd Module"), vec!["1"]); }
#[test] fn arith_exponentiation() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(2 ^ 3)\nEnd Sub\nEnd Module"), vec!["8"]); }
#[test] fn arith_negation() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10\nConsole.WriteLine(-x)\nEnd Sub\nEnd Module"), vec!["-10"]); }
#[test] fn arith_unary_plus() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = -10\nConsole.WriteLine(+x)\nEnd Sub\nEnd Module"), vec!["-10"]); }
#[test] fn arith_addition_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10\nx += 5\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["15"]); }

#[test] fn arith_subtraction_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10\nx -= 5\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn arith_multiplication_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10\nx *= 5\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["50"]); }
#[test] fn arith_division_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10.0\nx /= 4\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["2.5"]); }
#[test] fn arith_integer_division_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 10\nx \\= 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["3"]); }
#[test] fn arith_exponentiation_assignment() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 2.0\nx ^= 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"), vec!["8"]); }

#[test] fn arith_precedence_mult_add() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(2 + 3 * 4)\nEnd Sub\nEnd Module"), vec!["14"]); }
#[test] fn arith_precedence_exp_mult() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(2 * 3 ^ 2)\nEnd Sub\nEnd Module"), vec!["18"]); }
#[test] fn arith_precedence_parens() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine((2 + 3) * 4)\nEnd Sub\nEnd Module"), vec!["20"]); }
#[test] fn arith_division_by_zero_double() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 1.0\nConsole.WriteLine(x / 0.0 > 0)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn arith_integer_division_by_zero_throws() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 1\nDim y = 0\nTry\nDim z = x \\ y\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"), vec!["Caught"]); }

#[test] fn arith_overflow_integer() { assert_eq!(run_vb("Module M\nSub Main()\nDim x = 2147483647\nTry\nDim z = x + 1\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"), vec!["Caught"]); }
#[test] fn arith_modulo_floating_point() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(5.5 Mod 2.1)\nEnd Sub\nEnd Module"), vec!["1.3"]); }
#[test] fn arith_modulo_negative() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(-5 Mod 3)\nEnd Sub\nEnd Module"), vec!["-2"]); }
#[test] fn arith_division_integer_rounding() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(11 \\ 3)\nEnd Sub\nEnd Module"), vec!["3"]); } // Integer division truncates towards zero
#[test] fn arith_division_integer_rounding_negative() { assert_eq!(run_vb("Module M\nSub Main()\nConsole.WriteLine(-11 \\ 3)\nEnd Sub\nEnd Module"), vec!["-3"]); }
