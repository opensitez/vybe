use super::helpers::run_vb;

#[test] fn inherit_basic() { assert_eq!(run_vb("Class B\nPublic V As Integer = 1\nEnd Class\nClass C\nInherits B\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"), vec!["1"]); }
#[test] fn inherit_mybase_method() { assert_eq!(run_vb("Class B\nPublic Function GetV() As Integer\nReturn 10\nEnd Function\nEnd Class\nClass C\nInherits B\nPublic Function GetBaseV() As Integer\nReturn MyBase.GetV()\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetBaseV())\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn inherit_mybase_property() { assert_eq!(run_vb("Class B\nPublic Property V As Integer = 5\nEnd Class\nClass C\nInherits B\nPublic Function GetBaseV() As Integer\nReturn MyBase.V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetBaseV())\nEnd Sub\nEnd Module"), vec!["5"]); }
#[test] fn inherit_mybase_constructor() { assert_eq!(run_vb("Class B\nPublic V As Integer\nPublic Sub New(val As Integer)\nV = val\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Sub New()\nMyBase.New(20)\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"), vec!["20"]); }
#[test] fn inherit_mybase_not_allowed_in_module() { assert_eq!(run_vb("Module M\n' MyBase.ToString() ' Modules don't inherit from Object directly like classes in terms of MyBase\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }

#[test] fn inherit_myclass_method() { assert_eq!(run_vb("Class B\nPublic Overridable Function GetV() As Integer\nReturn 1\nEnd Function\nPublic Function CallMyClass() As Integer\nReturn MyClass.GetV()\nEnd Function\nEnd Class\nClass C\nInherits B\nPublic Overrides Function GetV() As Integer\nReturn 2\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.CallMyClass())\nEnd Sub\nEnd Module"), vec!["1"]); } // MyClass bypasses override
#[test] fn inherit_me_method() { assert_eq!(run_vb("Class B\nPublic Overridable Function GetV() As Integer\nReturn 1\nEnd Function\nPublic Function CallMe() As Integer\nReturn Me.GetV()\nEnd Function\nEnd Class\nClass C\nInherits B\nPublic Overrides Function GetV() As Integer\nReturn 2\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.CallMe())\nEnd Sub\nEnd Module"), vec!["2"]); } // Me respects override
#[test] fn inherit_multiple_levels() { assert_eq!(run_vb("Class A\nPublic V As Integer = 1\nEnd Class\nClass B\nInherits A\nEnd Class\nClass C\nInherits B\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"), vec!["1"]); }
#[test] fn inherit_multiple_classes_fails() { assert_eq!(run_vb("Class A\nEnd Class\nClass B\nEnd Class\n' Class C: Inherits A, B ' VB only supports single inheritance\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn inherit_structure_fails() { assert_eq!(run_vb("Structure S\nEnd Structure\n' Class C: Inherits S ' Cannot inherit from structure\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }

#[test]
fn graphics_drawline_sequence_runs_vb() {
    // VB benefits from the same drawing migration as C#: Graphics/Pen have no
    // ctor global, `g.DrawLine(...)` resolves through the component descriptor
    // and lowers inline. `Dim g As Graphics` types the receiver.
    let out = run_vb(
        "Imports System.Drawing\n\
         Imports System.Windows.Forms\n\
         Module M\nSub Main()\n\
         Dim g As Graphics = New PictureBox().CreateGraphics()\n\
         Dim p As New Pen(Color.Red, 2)\n\
         g.DrawLine(p, 0, 0, 10, 10)\n\
         Console.WriteLine(\"drew\")\n\
         End Sub\nEnd Module",
    );
    assert_eq!(out, vec!["drew"]);
}

#[test]
fn form_subclass_constructs_via_gui_host_after_ctor_removal() {
    // `class MyForm : Form` base construction routes through the shared
    // `try_emit_framework_control_base` (VB uses the same normalize_class →
    // emit_class_from_ast → compile_class path as C#). Control leaves no
    // longer emit a per-class ctor global; the host `new_Form` builds the
    // instance and inherited properties resolve through the descriptor.
    let out = run_vb(
        "Imports System.Windows.Forms\n\
         Public Class MyForm\nInherits Form\nEnd Class\n\
         Module M\nSub Main()\n\
         Dim f As New MyForm()\n\
         f.Text = \"hello\"\n\
         Console.WriteLine(f.Text)\n\
         Console.WriteLine(f.__control_type)\n\
         End Sub\nEnd Module",
    );
    assert_eq!(out, vec!["hello", "Form"]);
}

#[test] fn inherit_from_object_implicit() { assert_eq!(run_vb("Class C\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.ToString() IsNot Nothing)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn inherit_protected_member_access() { assert_eq!(run_vb("Class B\nProtected V As Integer = 42\nEnd Class\nClass C\nInherits B\nPublic Function GetV() As Integer\nReturn V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetV())\nEnd Sub\nEnd Module"), vec!["42"]); }
#[test] fn inherit_protected_member_no_access_outside() { assert_eq!(run_vb("Class B\nProtected V As Integer = 42\nEnd Class\nModule M\nSub Main()\nDim b1 As New B()\n' b1.V ' Fails\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn inherit_shadowing_field() { assert_eq!(run_vb("Class B\nPublic V As Integer = 1\nEnd Class\nClass C\nInherits B\nPublic Shadows V As Integer = 2\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"), vec!["2"]); }
#[test] fn inherit_shadowing_field_base_access() { assert_eq!(run_vb("Class B\nPublic V As Integer = 1\nEnd Class\nClass C\nInherits B\nPublic Shadows V As Integer = 2\nPublic Function GetBaseV() As Integer\nReturn MyBase.V\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.GetBaseV())\nEnd Sub\nEnd Module"), vec!["1"]); }

#[test] fn inherit_mustinherit_class_instantiation_fails() { assert_eq!(run_vb("MustInherit Class C\nEnd Class\nModule M\nSub Main()\n' Dim c1 As New C() ' Cannot instantiate MustInherit class\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn inherit_notinheritable_class_derivation_fails() { assert_eq!(run_vb("NotInheritable Class C\nEnd Class\n' Class D: Inherits C ' Cannot inherit from NotInheritable class\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn inherit_constructor_chaining_implicit() { assert_eq!(run_vb("Class B\nPublic V As Integer\nPublic Sub New()\nV = 10\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Sub New()\n' Implicit MyBase.New() runs\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(c1.V)\nEnd Sub\nEnd Module"), vec!["10"]); }
#[test] fn inherit_constructor_chaining_missing_parameterless_fails() { assert_eq!(run_vb("Class B\nPublic Sub New(v As Integer)\nEnd Sub\nEnd Class\nClass C\nInherits B\n' Public Sub New() ' Fails if it doesn't explicitly call MyBase.New(v) because B has no parameterless Sub New\nEnd Class\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"), vec!["Parsed"]); }
#[test] fn inherit_type_testing_typeof() { assert_eq!(run_vb("Class B\nEnd Class\nClass C\nInherits B\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nConsole.WriteLine(TypeOf c1 Is B)\nEnd Sub\nEnd Module"), vec!["True"]); }

#[test] fn inherit_type_conversion_directcast() { assert_eq!(run_vb("Class B\nEnd Class\nClass C\nInherits B\nEnd Class\nModule M\nSub Main()\nDim b1 As B = New C()\nDim c1 = DirectCast(b1, C)\nConsole.WriteLine(c1 IsNot Nothing)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn inherit_type_conversion_trycast() { assert_eq!(run_vb("Class B\nEnd Class\nClass C\nInherits B\nEnd Class\nModule M\nSub Main()\nDim b1 As New B()\nDim c1 = TryCast(b1, C)\nConsole.WriteLine(c1 Is Nothing)\nEnd Sub\nEnd Module"), vec!["True"]); }
#[test] fn inherit_virtual_call_in_constructor() { assert_eq!(run_vb("Class B\nPublic Sub New()\nM1()\nEnd Sub\nPublic Overridable Sub M1()\nConsole.WriteLine(\"B\")\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Overrides Sub M1()\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nEnd Sub\nEnd Module"), vec!["C"]); } // In VB, virtual calls in constructors reach derived implementations
#[test] fn inherit_shadowed_method_access_via_base_ref() { assert_eq!(run_vb("Class B\nPublic Sub M1()\nConsole.WriteLine(\"B\")\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Shadows Sub M1()\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim b1 As B = New C()\nb1.M1()\nEnd Sub\nEnd Module"), vec!["B"]); }
#[test] fn inherit_overridable_method_access_via_base_ref() { assert_eq!(run_vb("Class B\nPublic Overridable Sub M1()\nConsole.WriteLine(\"B\")\nEnd Sub\nEnd Class\nClass C\nInherits B\nPublic Overrides Sub M1()\nConsole.WriteLine(\"C\")\nEnd Sub\nEnd Class\nModule M\nSub Main()\nDim b1 As B = New C()\nb1.M1()\nEnd Sub\nEnd Module"), vec!["C"]); }
