use super::helpers::run_vb;

#[test]
fn enum_basic() {
    assert_eq!(
        run_vb(
            "Enum E\nA\nB\nEnd Enum\nModule M\nSub Main()\nDim e1 As E = E.B\nConsole.WriteLine(CInt(e1))\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
}
#[test]
fn enum_explicit_values() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 10\nB = 20\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.B))\nEnd Sub\nEnd Module"
        ),
        vec!["20"]
    );
}
#[test]
fn enum_mixed_values() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 5\nB\nC\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.C))\nEnd Sub\nEnd Module"
        ),
        vec!["7"]
    );
}
#[test]
fn enum_underlying_type_byte() {
    assert_eq!(
        run_vb(
            "Enum E As Byte\nA = 255\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(E.A.GetType().Name & CInt(E.A))\nEnd Sub\nEnd Module"
        ),
        vec!["E255"]
    );
}
#[test]
fn enum_underlying_type_long() {
    assert_eq!(
        run_vb(
            "Enum E As Long\nA = 3000000000L\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(E.A.GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["E"]
    );
}

#[test]
fn enum_tostring() {
    assert_eq!(
        run_vb(
            "Enum E\nFirstVal\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(E.FirstVal.ToString())\nEnd Sub\nEnd Module"
        ),
        vec!["FirstVal"]
    );
}
#[test]
fn enum_bitwise_flags_attribute_implicit() {
    assert_eq!(
        run_vb(
            "Enum Flags\nA = 1\nB = 2\nC = 4\nEnd Enum\nModule M\nSub Main()\nDim f = Flags.A Or Flags.B\nConsole.WriteLine(CInt(f))\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn enum_parse_string() {
    assert_eq!(
        run_vb(
            "Enum E\nValA\nEnd Enum\nModule M\nSub Main()\nDim e1 = CType([Enum].Parse(GetType(E), \"ValA\"), E)\nConsole.WriteLine(CInt(e1))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn enum_invalid_underlying_type() {
    assert_eq!(
        run_vb(
            "' Enum E As String ' Enums must be integral types\n' A = \"A\"\n' End Enum\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn enum_negative_values() {
    assert_eq!(
        run_vb(
            "Enum E\nA = -1\nB\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.B))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}

#[test]
fn enum_same_values() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 1\nB = 1\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(E.A = E.B)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn enum_scoping() {
    assert_eq!(
        run_vb(
            "Enum E\nA\nEnd Enum\nModule M\nSub Main()\n' Console.WriteLine(A) ' Requires E.A, not automatically imported into module scope unless implicitly supported in some legacy modes\nConsole.WriteLine(E.A.ToString())\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn enum_nested_in_class() {
    assert_eq!(
        run_vb(
            "Class C\nPublic Enum E\nA\nEnd Enum\nEnd Class\nModule M\nSub Main()\nConsole.WriteLine(CInt(C.E.A))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn enum_nested_in_structure() {
    assert_eq!(
        run_vb(
            "Structure S\nPublic Enum E\nA\nEnd Enum\nEnd Structure\nModule M\nSub Main()\nConsole.WriteLine(CInt(S.E.A))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn enum_forward_reference() {
    assert_eq!(
        run_vb(
            "Enum E\nA = B\nB = 5\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.A))\nEnd Sub\nEnd Module"
        ),
        vec!["5"]
    );
}

#[test]
fn enum_circular_reference_fails() {
    assert_eq!(
        run_vb(
            "Enum E\n' A = B\n' B = A\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn enum_math_operations() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 1\nB = 2\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.A + E.B))\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn enum_comparison() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 1\nB = 2\nEnd Enum\nModule M\nSub Main()\nConsole.WriteLine(E.A < E.B)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn enum_implicit_conversion_from_integer() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nEnum E\nA = 1\nEnd Enum\nModule M\nSub Main()\nDim e1 As E = 1\nConsole.WriteLine(e1.ToString())\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn enum_explicit_conversion_from_integer_strict_on() {
    assert_eq!(
        run_vb(
            "Option Strict On\nEnum E\nA = 1\nEnd Enum\nModule M\nSub Main()\nDim e1 As E = CType(1, E)\nConsole.WriteLine(e1.ToString())\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}

#[test]
fn enum_no_members() {
    assert_eq!(
        run_vb(
            "Enum E\nEnd Enum\nModule M\nSub Main()\nDim e1 As E\nConsole.WriteLine(CInt(e1))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn enum_shadowing() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 1\nEnd Enum\nClass C\nPublic Enum E\nA = 2\nEnd Enum\nEnd Class\nModule M\nSub Main()\nConsole.WriteLine(CInt(E.A) + CInt(C.E.A))\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn enum_as_property_type() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 5\nEnd Enum\nClass C\nPublic Property V As E\nEnd Class\nModule M\nSub Main()\nDim c1 As New C()\nc1.V = E.A\nConsole.WriteLine(CInt(c1.V))\nEnd Sub\nEnd Module"
        ),
        vec!["5"]
    );
}
#[test]
fn enum_as_method_return_type() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 10\nEnd Enum\nModule M\nFunction GetE() As E\nReturn E.A\nEnd Function\nSub Main()\nConsole.WriteLine(CInt(GetE()))\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn enum_default_value() {
    assert_eq!(
        run_vb(
            "Enum E\nA = 5\nEnd Enum\nModule M\nSub Main()\nDim e1 As E ' Default is 0, even if not defined in Enum\nConsole.WriteLine(CInt(e1))\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
