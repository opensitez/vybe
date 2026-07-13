use super::helpers::run_vb;

// Events and delegates
#[test] fn event_basic() { assert_eq!(run_vb(r#"Class C: Public Event E(): Public Sub DoE(): RaiseEvent E(): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): AddHandler obj.E, Sub() Console.WriteLine("Fired"): obj.DoE(): End Sub: End Module"#), vec!["Fired"]); }
#[test] fn event_with_args() { assert_eq!(run_vb(r#"Class C: Public Event E(x As Integer): Public Sub DoE(): RaiseEvent E(42): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): AddHandler obj.E, Sub(x) Console.WriteLine(x): obj.DoE(): End Sub: End Module"#), vec!["42"]); }
#[test] fn event_removehandler() { assert_eq!(run_vb(r#"Class C: Public Event E(): Public Sub DoE(): RaiseEvent E(): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): Dim h As System.Action = Sub() Console.WriteLine("X"): AddHandler obj.E, h: RemoveHandler obj.E, h: obj.DoE(): Console.WriteLine("Done"): End Sub: End Module"#), vec!["Done"]); }
#[test] fn event_custom_accessor() { assert_eq!(run_vb(r#"Class C: Private _h As System.Action: Public Custom Event E As System.Action: AddHandler(v As System.Action): _h = v: End AddHandler: RemoveHandler(v As System.Action): _h = Nothing: End RemoveHandler: RaiseEvent(): _h?(): End RaiseEvent: End Event: Public Sub DoE(): RaiseEvent E(): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): AddHandler obj.E, Sub() Console.WriteLine("Custom"): obj.DoE(): End Sub: End Module"#), vec!["Custom"]); }

// Attributes
#[test] fn attr_class() { assert_eq!(run_vb(r#"<System.Serializable> Class C: End Class: Module M: Sub Main(): Console.WriteLine(GetType(C).IsSerializable): End Sub: End Module"#), vec!["True"]); }
#[test] fn attr_method() { assert_eq!(run_vb(r#"Class C: <System.Obsolete("Old")> Public Sub M1(): End Sub: End Class: Module M: Sub Main(): Dim m = GetType(C).GetMethod("M1"): Console.WriteLine(m.GetCustomAttributes(False).Length > 0): End Sub: End Module"#), vec!["True"]); }
#[test] fn attr_multiple() { assert_eq!(run_vb(r#"<System.Serializable, System.Obsolete> Class C: End Class: Module M: Sub Main(): Console.WriteLine(GetType(C).GetCustomAttributes(False).Length >= 2): End Sub: End Module"#), vec!["True"]); }
#[test] fn attr_property() { assert_eq!(run_vb(r#"Class C: <System.Obsolete> Public Property P As Integer: End Class: Module M: Sub Main(): Dim p = GetType(C).GetProperty("P"): Console.WriteLine(p.GetCustomAttributes(False).Length > 0): End Sub: End Module"#), vec!["True"]); }
#[test] fn attr_target_assembly() { assert_eq!(run_vb(r#"<Assembly: System.Reflection.AssemblyTitle("Test")> Module M: Sub Main(): Console.WriteLine("Parsed"): End Sub: End Module"#), vec!["Parsed"]); }
#[test] fn attr_custom() { assert_eq!(run_vb(r#"<AttributeUsage(AttributeTargets.Class)> Class MyAttr: Inherits Attribute: Public Name As String: Public Sub New(n As String): Name = n: End Sub: End Class: <MyAttr("Test")> Class C: End Class: Module M: Sub Main(): Dim a = CType(GetType(C).GetCustomAttributes(False)(0), MyAttr): Console.WriteLine(a.Name): End Sub: End Module"#), vec!["Test"]); }

// Delegates
#[test] fn delegate_basic() { assert_eq!(run_vb(r#"Delegate Sub D(): Module M: Sub Main(): Dim act As D = Sub() Console.WriteLine("Del"): act(): End Sub: End Module"#), vec!["Del"]); }
#[test] fn delegate_func() { assert_eq!(run_vb(r#"Delegate Function D(x As Integer) As Integer: Module M: Sub Main(): Dim fn As D = Function(x) x * 2: Console.WriteLine(fn(21)): End Sub: End Module"#), vec!["42"]); }
#[test] fn delegate_invoke() { assert_eq!(run_vb(r#"Delegate Sub D(): Module M: Sub Main(): Dim act As D = Sub() Console.WriteLine("Inv"): act.Invoke(): End Sub: End Module"#), vec!["Inv"]); }
#[test] fn delegate_multicast() { assert_eq!(run_vb(r#"Delegate Sub D(): Module M: Sub Main(): Dim a As D = Sub() Console.WriteLine("A"): Dim b As D = Sub() Console.WriteLine("B"): Dim c = CType(System.Delegate.Combine(a, b), D): c(): End Sub: End Module"#), vec!["A", "B"]); }
#[test] fn delegate_relaxed() { assert_eq!(run_vb(r#"Delegate Sub D(): Module M: Sub Target(x As Integer): Console.WriteLine(x): End Sub: Sub Main(): ' Relaxed delegate instantiation (dropping args) is a VB feature: Dim act As D = AddressOf Target: act(): End Sub: End Module"#), vec!["0"]); }

// Generics
#[test] fn generic_class() { assert_eq!(run_vb(r#"Class C(Of T): Public V As T: End Class: Module M: Sub Main(): Dim obj As New C(Of String)(): obj.V = "A": Console.WriteLine(obj.V): End Sub: End Module"#), vec!["A"]); }
#[test] fn generic_method() { assert_eq!(run_vb(r#"Module M: Function Id(Of T)(v As T) As T: Return v: End Function: Sub Main(): Console.WriteLine(Id(42)): End Sub: End Module"#), vec!["42"]); }
#[test] fn generic_constraint_class() { assert_eq!(run_vb(r#"Class C(Of T As Class): End Class: Module M: Sub Main(): Console.WriteLine("Parsed"): End Sub: End Module"#), vec!["Parsed"]); }
#[test] fn generic_constraint_struct() { assert_eq!(run_vb(r#"Class C(Of T As Structure): End Class: Module M: Sub Main(): Console.WriteLine("Parsed"): End Sub: End Module"#), vec!["Parsed"]); }
#[test] fn generic_constraint_new() { assert_eq!(run_vb(r#"Class C(Of T As New): Public Function Create() As T: Return New T(): End Function: End Class: Class Item: Public Sub New(): Console.WriteLine("New"): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(Of Item)(): obj.Create(): End Sub: End Module"#), vec!["New"]); }

// Properties
#[test] fn prop_auto() { assert_eq!(run_vb(r#"Class C: Public Property P As Integer: End Class: Module M: Sub Main(): Dim obj As New C(): obj.P = 10: Console.WriteLine(obj.P): End Sub: End Module"#), vec!["10"]); }
#[test] fn prop_auto_init() { assert_eq!(run_vb(r#"Class C: Public Property P As Integer = 42: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.P): End Sub: End Module"#), vec!["42"]); }
#[test] fn prop_readonly() { assert_eq!(run_vb(r#"Class C: Public ReadOnly Property P As Integer = 10: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.P): End Sub: End Module"#), vec!["10"]); }
#[test] fn prop_writeonly() { assert_eq!(run_vb(r#"Class C: Private _p As Integer: Public WriteOnly Property P As Integer: Set(v As Integer): _p = v: End Set: End Property: Public Function GetP() As Integer: Return _p: End Function: End Class: Module M: Sub Main(): Dim obj As New C(): obj.P = 20: Console.WriteLine(obj.GetP()): End Sub: End Module"#), vec!["20"]); }
#[test] fn prop_indexed() { assert_eq!(run_vb(r#"Class C: Public Property P(i As Integer) As Integer: Get: Return i * 2: End Get: Set(v As Integer): End Set: End Property: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.P(5)): End Sub: End Module"#), vec!["10"]); }

// Inheritance
#[test] fn inherit_basic() { assert_eq!(run_vb(r#"Class B: Public V As Integer = 1: End Class: Class C: Inherits B: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.V): End Sub: End Module"#), vec!["1"]); }
#[test] fn inherit_override() { assert_eq!(run_vb(r#"Class B: Public Overridable Function M() As String: Return "B": End Function: End Class: Class C: Inherits B: Public Overrides Function M() As String: Return "C": End Function: End Class: Module M: Sub Main(): Dim obj As B = New C(): Console.WriteLine(obj.M()): End Sub: End Module"#), vec!["C"]); }
#[test] fn inherit_shadows() { assert_eq!(run_vb(r#"Class B: Public Function M() As String: Return "B": End Function: End Class: Class C: Inherits B: Public Shadows Function M() As String: Return "C": End Function: End Class: Module M: Sub Main(): Dim obj1 As New C(): Dim obj2 As B = obj1: Console.WriteLine(obj1.M() & obj2.M()): End Sub: End Module"#), vec!["CB"]); }
#[test] fn inherit_mybase() { assert_eq!(run_vb(r#"Class B: Public Overridable Function M() As String: Return "B": End Function: End Class: Class C: Inherits B: Public Overrides Function M() As String: Return MyBase.M() & "C": End Function: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.M()): End Sub: End Module"#), vec!["BC"]); }
#[test] fn inherit_myclass() { assert_eq!(run_vb(r#"Class B: Public Overridable Function M() As String: Return "B": End Function: Public Function Test() As String: Return MyClass.M(): End Function: End Class: Class C: Inherits B: Public Overrides Function M() As String: Return "C": End Function: End Class: Module M: Sub Main(): Dim obj As New C(): Console.WriteLine(obj.Test()): End Sub: End Module"#), vec!["B"]); }

// Interfaces
#[test] fn interface_basic() { assert_eq!(run_vb(r#"Interface I: Sub Test(): End Interface: Class C: Implements I: Public Sub Test() Implements I.Test: Console.WriteLine("I"): End Sub: End Class: Module M: Sub Main(): Dim obj As I = New C(): obj.Test(): End Sub: End Module"#), vec!["I"]); }
#[test] fn interface_multiple() { assert_eq!(run_vb(r#"Interface I1: Sub Test(): End Interface: Interface I2: Sub Test(): End Interface: Class C: Implements I1, I2: Public Sub Test() Implements I1.Test, I2.Test: Console.WriteLine("B"): End Sub: End Class: Module M: Sub Main(): Dim obj As New C(): obj.Test(): End Sub: End Module"#), vec!["B"]); }
#[test] fn interface_inherit() { assert_eq!(run_vb(r#"Interface I1: Sub T1(): End Interface: Interface I2: Inherits I1: Sub T2(): End Interface: Class C: Implements I2: Public Sub T1() Implements I2.T1: End Sub: Public Sub T2() Implements I2.T2: End Sub: End Class: Module M: Sub Main(): Console.WriteLine("Parsed"): End Sub: End Module"#), vec!["Parsed"]); }

// Structures
#[test] fn struct_basic() { assert_eq!(run_vb(r#"Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim s1 As New S(): s1.V = 1: Console.WriteLine(s1.V): End Sub: End Module"#), vec!["1"]); }
#[test] fn struct_copy() { assert_eq!(run_vb(r#"Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim s1 As New S(): s1.V = 1: Dim s2 = s1: s2.V = 2: Console.WriteLine(s1.V): End Sub: End Module"#), vec!["1"]); }

// Enums
#[test] fn enum_basic() { assert_eq!(run_vb(r#"Enum E: A: B: C: End Enum: Module M: Sub Main(): Console.WriteLine(E.B): End Sub: End Module"#), vec!["1"]); }
#[test] fn enum_explicit() { assert_eq!(run_vb(r#"Enum E: A = 10: B = 20: End Enum: Module M: Sub Main(): Console.WriteLine(E.B): End Sub: End Module"#), vec!["20"]); }
#[test] fn enum_flags() { assert_eq!(run_vb(r#"<System.Flags> Enum E: A = 1: B = 2: C = 4: End Enum: Module M: Sub Main(): Dim val = E.A Or E.C: Console.WriteLine(val): End Sub: End Module"#), vec!["5"]); }

// Overloading
#[test] fn overload_basic() { assert_eq!(run_vb(r#"Module M: Sub Test(x As Integer): Console.WriteLine("I"): End Sub: Sub Test(x As String): Console.WriteLine("S"): End Sub: Sub Main(): Test("A"): End Sub: End Module"#), vec!["S"]); }
#[test] fn overload_byref() { assert_eq!(run_vb(r#"Module M: Sub Test(ByRef x As Integer): Console.WriteLine("ByRef"): End Sub: Sub Test(x As String): End Sub: Sub Main(): Dim y As Integer = 1: Test(y): End Sub: End Module"#), vec!["ByRef"]); }

// Operator Overloading
#[test] fn op_overload_add() { assert_eq!(run_vb(r#"Class C: Public V As Integer: Public Shared Operator +(a As C, b As C) As C: Return New C With {.V = a.V + b.V}: End Operator: End Class: Module M: Sub Main(): Dim c1 As New C With {.V = 1}: Dim c2 As New C With {.V = 2}: Console.WriteLine((c1 + c2).V): End Sub: End Module"#), vec!["3"]); }
#[test] fn op_overload_eq() { assert_eq!(run_vb(r#"Class C: Public V As Integer: Public Shared Operator =(a As C, b As C) As Boolean: Return a.V = b.V: End Operator: Public Shared Operator <>(a As C, b As C) As Boolean: Return a.V <> b.V: End Operator: End Class: Module M: Sub Main(): Dim c1 As New C With {.V = 1}: Dim c2 As New C With {.V = 1}: Console.WriteLine(c1 = c2): End Sub: End Module"#), vec!["True"]); }

// Extension Methods
#[test] fn ext_method_string() { assert_eq!(run_vb(r#"Imports System.Runtime.CompilerServices: Module Ext: <Extension()> Public Function Excite(s As String) As String: Return s & "!": End Function: End Module: Module M: Sub Main(): Console.WriteLine("A".Excite()): End Sub: End Module"#), vec!["A!"]); }
#[test] fn ext_method_generic() { assert_eq!(run_vb(r#"Imports System.Runtime.CompilerServices: Module Ext: <Extension()> Public Function Wrap(Of T)(val As T) As String: Return "[" & val.ToString() & "]": End Function: End Module: Module M: Sub Main(): Console.WriteLine((42).Wrap()): End Sub: End Module"#), vec!["[42]"]); }

// Namespaces
#[test] fn namespace_basic() { assert_eq!(run_vb(r#"Namespace NS: Class C: Public V As Integer = 10: End Class: End Namespace: Module M: Sub Main(): Dim obj As New NS.C(): Console.WriteLine(obj.V): End Sub: End Module"#), vec!["10"]); }
#[test] fn namespace_nested() { assert_eq!(run_vb(r#"Namespace NS1.NS2: Class C: Public V As Integer = 20: End Class: End Namespace: Module M: Sub Main(): Dim obj As New NS1.NS2.C(): Console.WriteLine(obj.V): End Sub: End Module"#), vec!["20"]); }

// Module features
#[test] fn module_shared_implicit() { assert_eq!(run_vb(r#"Module Data: Public V As Integer = 10: End Module: Module M: Sub Main(): Console.WriteLine(Data.V): End Sub: End Module"#), vec!["10"]); }
#[test] fn module_aliasing() { assert_eq!(run_vb(r#"Imports Alias = System.Console: Module M: Sub Main(): Alias.WriteLine("Alias"): End Sub: End Module"#), vec!["Alias"]); }
