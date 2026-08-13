use super::helpers::run_vb;

#[test]
fn object_implicit_declaration() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim x\nx = 10\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Int32"]
    );
}
#[test]
fn object_reassignment_type_change() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim x\nx = 10\nx = \"Hello\"\nConsole.WriteLine(x.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["String"]
    );
}
#[test]
fn object_array_mixed_types() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Object = {1, \"A\", True}\nConsole.WriteLine(arr(1).GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["String"]
    );
}
#[test]
fn late_binding_method_call() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Function M1() As String\nReturn \"OK\"\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nConsole.WriteLine(obj.M1())\nEnd Sub\nEnd Module"
        ),
        vec!["OK"]
    );
}
/// `Dim obj As Object` and its first assignment as SEPARATE statements — the
/// other half of the spelling every test above uses inline. Ground truth
/// (dotnet 10): `OK`.
#[test]
fn late_binding_split_declaration_and_assignment() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Function M1() As String\nReturn \"OK\"\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim obj As Object\nobj = New C()\nConsole.WriteLine(obj.M1())\nEnd Sub\nEnd Module"
        ),
        vec!["OK"]
    );
}

/// The same split spelling against a PLATFORM type, whose members are
/// compile-time emits with nothing on the object to dispatch to. This is the
/// case that broke: `Append` is also a LINQ name, so an untyped receiver sent
/// the call to LINQ `Append` over an iterable and it appended nothing —
/// silently, with `ToString()` answering `[object StringBuilder]`.
/// Ground truth (dotnet 10): `built`.
#[test]
fn late_binding_split_declaration_platform_type() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim sb As Object\nsb = New System.Text.StringBuilder()\nsb.Append(\"built\")\nConsole.WriteLine(sb.ToString())\nEnd Sub\nEnd Module"
        ),
        vec!["built"]
    );
}

#[test]
fn late_binding_property_get() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Property P As Integer = 42\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nConsole.WriteLine(obj.P)\nEnd Sub\nEnd Module"
        ),
        vec!["42"]
    );
}

#[test]
fn late_binding_property_set() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Property P As Integer\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nobj.P = 10\nConsole.WriteLine(obj.P)\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn late_binding_missing_method_fails() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nTry\nobj.Missing()\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn late_binding_implicit_coercion_args() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Function Add(a As Integer, b As Integer) As Integer\nReturn a + b\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nConsole.WriteLine(obj.Add(\"1\", \"2\"))\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
} // "1" and "2" implicitly converted to Integer during late bound call
#[test]
fn late_binding_overload_resolution() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic Function M(a As Integer) As String\nReturn \"Int\"\nEnd Function\nPublic Function M(a As String) As String\nReturn \"Str\"\nEnd Function\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nConsole.WriteLine(obj.M(10))\nEnd Sub\nEnd Module"
        ),
        vec!["Int"]
    );
}
#[test]
fn object_default_property() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nDefault Public Property Item(i As Integer) As String\nGet\nReturn \"Item\" & i\nEnd Get\nSet(v As String)\nEnd Set\nEnd Property\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nConsole.WriteLine(obj(1))\nEnd Sub\nEnd Module"
        ),
        vec!["Item1"]
    );
}

#[test]
fn object_is_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj1 As New Object()\nDim obj2 = obj1\nConsole.WriteLine(obj1 Is obj2)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn object_isnot_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj1 As New Object()\nDim obj2 As New Object()\nConsole.WriteLine(obj1 IsNot obj2)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn object_equality_operator() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim obj1 As Object = 1\nDim obj2 As Object = 1\nConsole.WriteLine(obj1 = obj2)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
} // By-value comparison when underlying type supports it in late binding
#[test]
fn object_equality_reference_types() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nEnd Class\nModule M\nSub Main()\nDim obj1 As Object = New C()\nDim obj2 As Object = New C()\nTry\nDim r = (obj1 = obj2)\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
} // Late binding = on two objects without Operator = throws

#[test]
fn variant_type_function() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = 10\nConsole.WriteLine(VarType(obj))\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
} // vbInteger
#[test]
fn typeof_is_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = \"A\"\nConsole.WriteLine(TypeOf obj Is String)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn typeof_isnot_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = 10\nConsole.WriteLine(TypeOf obj IsNot String)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn object_nothing_assignment() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = Nothing\nConsole.WriteLine(obj Is Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn object_gettype_method() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = True\nConsole.WriteLine(obj.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Boolean"]
    );
}

#[test]
fn late_binding_with_block() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nClass C\nPublic V As Integer = 10\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nWith obj\n.V = 20\nConsole.WriteLine(.V)\nEnd With\nEnd Sub\nEnd Module"
        ),
        vec!["20"]
    );
}
#[test]
fn late_binding_array_access() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim arr() As Integer = {1, 2, 3}\nDim obj As Object = arr\nConsole.WriteLine(obj(1))\nEnd Sub\nEnd Module"
        ),
        vec!["2"]
    );
}
#[test]
fn late_binding_collection_add() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim col As Object = New Collection()\ncol.Add(\"A\")\nConsole.WriteLine(col.Count)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn late_binding_for_each() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim arr() As Integer = {10}\nDim obj As Object = arr\nFor Each v In obj\nConsole.WriteLine(v)\nNext\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn late_binding_math_method() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim obj As Object = 4.5\nConsole.WriteLine(Math.Round(obj))\nEnd Sub\nEnd Module"
        ),
        vec!["4"]
    );
}

#[test]
fn late_binding_option_strict_on_fails() {
    assert_eq!(
        run_vb(
            "Option Strict On\nClass C\nPublic V As Integer = 1\nEnd Class\nModule M\nSub Main()\n' Dim obj As Object = New C(): Console.WriteLine(obj.V) ' Compiler error in Option Strict On\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
