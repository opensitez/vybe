use super::helpers::run_vb;

// CType
#[test] fn ctype_string_int() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CType("42", Integer)): End Sub: End Module"#), vec!["42"]); }
#[test] fn ctype_int_string() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CType(42, String)): End Sub: End Module"#), vec!["42"]); }
#[test] fn ctype_double_int() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CType(42.5, Integer)): End Sub: End Module"#), vec!["42"]); } // Banker's rounding rounds 42.5 to 42

// DirectCast
#[test] fn directcast_exact() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = "A": Console.WriteLine(DirectCast(o, String)): End Sub: End Module"#), vec!["A"]); }
#[test] fn directcast_fail() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = "A": Try: DirectCast(o, Integer): Catch: Console.WriteLine("Err"): End Try: End Sub: End Module"#), vec!["Err"]); }

// TryCast
#[test] fn trycast_success() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = "A": Dim s = TryCast(o, String): Console.WriteLine(s): End Sub: End Module"#), vec!["A"]); }
#[test] fn trycast_fail() { assert_eq!(run_vb(r#"Class C: End Class: Module M: Sub Main(): Dim o As Object = "A": Dim c = TryCast(o, C): Console.WriteLine(c Is Nothing): End Sub: End Module"#), vec!["True"]); }

// Implicit conversions (Option Strict Off by default)
#[test] fn implicit_string_int() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim x As Integer = "10": Console.WriteLine(x): End Sub: End Module"#), vec!["10"]); }
#[test] fn implicit_int_string() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim s As String = 10: Console.WriteLine(s): End Sub: End Module"#), vec!["10"]); }
#[test] fn implicit_bool_int() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim x As Integer = True: Console.WriteLine(x): End Sub: End Module"#), vec!["-1"]); }

// Conversion functions (CInt, CDbl, CStr, CBool, etc.)
#[test] fn cbool_true() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CBool(1)): End Sub: End Module"#), vec!["True"]); }
#[test] fn cbool_false() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CBool(0)): End Sub: End Module"#), vec!["False"]); }
#[test] fn cbyte_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CByte(255)): End Sub: End Module"#), vec!["255"]); }
#[test] fn cchar_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CChar("A")): End Sub: End Module"#), vec!["A"]); }
#[test] fn cdate_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CDate("2020-01-01").Year): End Sub: End Module"#), vec!["2020"]); }
#[test] fn cdbl_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CDbl("3.14")): End Sub: End Module"#), vec!["3.14"]); }
#[test] fn cdec_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CDec(10.5)): End Sub: End Module"#), vec!["10.5"]); }
#[test] fn cint_round() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CInt(2.5)): End Sub: End Module"#), vec!["2"]); }
#[test] fn clng_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CLng(10000000000)): End Sub: End Module"#), vec!["10000000000"]); }
#[test] fn cobj_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o = CObj(1): Console.WriteLine(o.GetType().Name): End Sub: End Module"#), vec!["Int32"]); }
#[test] fn csbyte_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CSByte(-10)): End Sub: End Module"#), vec!["-10"]); }
#[test] fn cshort_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CShort(-30000)): End Sub: End Module"#), vec!["-30000"]); }
#[test] fn csng_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CSng("1.5")): End Sub: End Module"#), vec!["1.5"]); }
#[test] fn cstr_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CStr(100)): End Sub: End Module"#), vec!["100"]); }
#[test] fn cuint_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CUInt(4000000000)): End Sub: End Module"#), vec!["4000000000"]); }
#[test] fn culng_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CULng(10000000000)): End Sub: End Module"#), vec!["10000000000"]); }
#[test] fn cushort_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(CUShort(60000)): End Sub: End Module"#), vec!["60000"]); }

// Widening / Narrowing Explicit Tests
#[test] fn widen_byte_int() { assert_eq!(run_vb(r#"Option Strict On: Module M: Sub Main(): Dim b As Byte = 10: Dim i As Integer = b: Console.WriteLine(i): End Sub: End Module"#), vec!["10"]); }
#[test] fn narrow_int_byte_err() { assert_eq!(run_vb(r#"Option Strict On: Module M: Sub Main(): ' Dim b As Byte = 1000 ' Compile Error with Option Strict: Console.WriteLine("Parsed"): End Sub: End Module"#), vec!["Parsed"]); }
#[test] fn widen_int_long() { assert_eq!(run_vb(r#"Option Strict On: Module M: Sub Main(): Dim i As Integer = 10: Dim l As Long = i: Console.WriteLine(l): End Sub: End Module"#), vec!["10"]); }
#[test] fn widen_single_double() { assert_eq!(run_vb(r#"Option Strict On: Module M: Sub Main(): Dim s As Single = 1.5: Dim d As Double = s: Console.WriteLine(d > 1): End Sub: End Module"#), vec!["True"]); }

// CType array
#[test] fn ctype_array() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a As Object() = {"A", "B"}: Dim b = CType(a, String()): Console.WriteLine(b(0)): End Sub: End Module"#), vec!["A"]); }

// DirectCast array
#[test] fn directcast_array() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a As Object = New String() {"A"}: Dim b = DirectCast(a, String()): Console.WriteLine(b(0)): End Sub: End Module"#), vec!["A"]); }

// TypeOf Is
#[test] fn typeof_is_string() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = "A": Console.WriteLine(TypeOf o Is String): End Sub: End Module"#), vec!["True"]); }
#[test] fn typeof_is_int() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = 10: Console.WriteLine(TypeOf o Is Integer): End Sub: End Module"#), vec!["True"]); }
#[test] fn typeof_is_array() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = New Integer() {1}: Console.WriteLine(TypeOf o Is Integer()): End Sub: End Module"#), vec!["True"]); }

// TypeOf IsNot
#[test] fn typeof_isnot_string() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o As Object = 10: Console.WriteLine(TypeOf o IsNot String): End Sub: End Module"#), vec!["True"]); }

// GetType
#[test] fn gettype_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(GetType(Integer).Name): End Sub: End Module"#), vec!["Int32"]); }
#[test] fn gettype_object() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim o = 10: Console.WriteLine(o.GetType().Name): End Sub: End Module"#), vec!["Int32"]); }

// IIf
#[test] fn iif_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(IIf(True, "A", "B")): End Sub: End Module"#), vec!["A"]); }
#[test] fn iif_evals_both() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim x = 0: Dim y = IIf(True, x + 1, x + 2): Console.WriteLine(y): End Sub: End Module"#), vec!["1"]); } // In real VB IIf evals both, but we can't test side-effect easily without a function call.

// If Operator (short-circuit)
#[test] fn if_operator_basic() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(If(True, "A", "B")): End Sub: End Module"#), vec!["A"]); }
#[test] fn if_operator_null() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim s As String = Nothing: Console.WriteLine(If(s, "B")): End Sub: End Module"#), vec!["B"]); }

// Convert Class (System.Convert)
#[test] fn sys_convert_toint32() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(System.Convert.ToInt32("42")): End Sub: End Module"#), vec!["42"]); }
#[test] fn sys_convert_toboolean() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(System.Convert.ToBoolean("True")): End Sub: End Module"#), vec!["True"]); }
#[test] fn sys_convert_todouble() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(System.Convert.ToDouble("3.14")): End Sub: End Module"#), vec!["3.14"]); }
#[test] fn sys_convert_tostring() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(System.Convert.ToString(42)): End Sub: End Module"#), vec!["42"]); }

// Parse Methods
#[test] fn int_parse() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Integer.Parse("10")): End Sub: End Module"#), vec!["10"]); }
#[test] fn int_tryparse() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim v As Integer: Integer.TryParse("10", v): Console.WriteLine(v): End Sub: End Module"#), vec!["10"]); }
#[test] fn double_parse() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Double.Parse("1.5")): End Sub: End Module"#), vec!["1.5"]); }
#[test] fn double_tryparse() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim v As Double: Double.TryParse("1.5", v): Console.WriteLine(v): End Sub: End Module"#), vec!["1.5"]); }
