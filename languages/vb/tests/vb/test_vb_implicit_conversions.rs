use super::helpers::run_vb;

#[test] fn implicit_widening_short_to_integer() { assert_eq!(run_vb("Module M\nSub Main()\nDim s As Short = 10\nDim i As Integer = s\nConsole.WriteLine(i.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn implicit_widening_integer_to_long() { assert_eq!(run_vb("Module M\nSub Main()\nDim i As Integer = 10\nDim l As Long = i\nConsole.WriteLine(l.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn implicit_widening_single_to_double() { assert_eq!(run_vb("Module M\nSub Main()\nDim s As Single = 1.5F\nDim d As Double = s\nConsole.WriteLine(d.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Double"]); }
#[test] fn implicit_widening_integer_to_decimal() { assert_eq!(run_vb("Module M\nSub Main()\nDim i As Integer = 10\nDim d As Decimal = i\nConsole.WriteLine(d.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Decimal"]); }
#[test] fn implicit_widening_char_to_string() { assert_eq!(run_vb("Module M\nSub Main()\nDim c As Char = \"A\"c\nDim s As String = c\nConsole.WriteLine(s.GetType().Name)\nEnd Sub\nEnd Module"), vec!["String"]); }

#[test] fn implicit_narrowing_long_to_integer() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim l As Long = 10\nDim i As Integer = l\nConsole.WriteLine(i.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn implicit_narrowing_double_to_single() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Double = 1.5\nDim s As Single = d\nConsole.WriteLine(s.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Single"]); }
#[test] fn implicit_narrowing_decimal_to_integer() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Decimal = 10D\nDim i As Integer = d\nConsole.WriteLine(i.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn implicit_narrowing_string_to_char_array() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim s As String = \"Hello\"\nDim arr() As Char = s.ToCharArray()\nConsole.WriteLine(arr.Length)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn implicit_narrowing_string_to_char() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim s As String = \"A\"\nDim c As Char = s\nConsole.WriteLine(c.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Char"]); }

#[test] fn implicit_string_to_integer_numeric() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim s As String = \"10\"\nDim i As Integer = s\nConsole.WriteLine(i.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int32"]); }
#[test] fn implicit_integer_to_string() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim i As Integer = 10\nDim s As String = i\nConsole.WriteLine(s.GetType().Name)\nEnd Sub\nEnd Module"), vec!["String"]); }
#[test] fn implicit_boolean_to_integer() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim b As Boolean = True\nDim i As Integer = b\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"), vec!["-1"]); }
#[test] fn implicit_integer_to_boolean() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim i As Integer = -1\nDim b As Boolean = i\nConsole.WriteLine(b)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn implicit_double_to_integer_bankers_rounding() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Double = 2.5\nDim i As Integer = d\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"), vec!["2"]); } // Bankers rounding rounds 2.5 to 2

#[test] fn implicit_double_to_integer_bankers_rounding_up() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Double = 3.5\nDim i As Integer = d\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"), vec!["4"]); } // Bankers rounding rounds 3.5 to 4
#[test] fn implicit_nothing_to_string() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim s As String = Nothing\nConsole.WriteLine(s = \"\")\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn implicit_nothing_to_integer() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim i As Integer = Nothing\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"), vec!["0"]); }
#[test] fn implicit_nothing_to_boolean() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim b As Boolean = Nothing\nConsole.WriteLine(b)\nEnd Sub\nEnd Module"), vec!["False"]); }
#[test] fn implicit_date_to_string() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim d As Date = #1/1/2000#\nDim s As String = d\nConsole.WriteLine(s.Contains(\"2000\"))\nEnd Sub\nEnd Module"), vec!["True"]); }

#[test] fn implicit_option_strict_on_widening_only() { assert_eq!(run_vb("Option Strict On\nModule M\nSub Main()\nDim i As Integer = 10\nDim l As Long = i\nConsole.WriteLine(l.GetType().Name)\nEnd Sub\nEnd Module"), vec!["Int64"]); }
#[test] fn implicit_option_strict_on_narrowing_fails() { assert_eq!(run_vb("Option Strict On\nModule M\nSub Main()\n' Dim l As Long = 10: Dim i As Integer = l ' Fails in Option Strict On\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn implicit_array_covariance_fails() { assert_eq!(run_vb("Option Strict Off\nModule M\nSub Main()\nDim strArr() As String = {\"A\", \"B\"}\n' Arrays are not covariant in VB like C#, an object array cannot just point to a string array\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn implicit_enum_to_integer() { assert_eq!(run_vb("Enum E\nVal = 5\nEnd Enum\nModule M\nSub Main()\nDim i As Integer = E.Val\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn implicit_integer_to_enum() { assert_eq!(run_vb("Option Strict Off\nEnum E\nVal = 5\nEnd Enum\nModule M\nSub Main()\nDim e As E = 5\nConsole.WriteLine(e.ToString())\nEnd Sub\nEnd Module"), vec!["Val"]); }
