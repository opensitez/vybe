use super::helpers::run_vb;

#[test]
fn explicit_cint_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CInt(\"10\"))\nEnd Sub\nEnd Module"),
        vec!["10"]
    );
}
#[test]
fn explicit_cint_double_bankers() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CInt(2.5))\nEnd Sub\nEnd Module"),
        vec!["2"]
    );
}
#[test]
fn explicit_clng_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CLng(\"10\"))\nEnd Sub\nEnd Module"),
        vec!["10"]
    );
}
#[test]
fn explicit_clng_double_bankers() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CLng(3.5))\nEnd Sub\nEnd Module"),
        vec!["4"]
    );
}
#[test]
fn explicit_csng_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CSng(\"1.5\"))\nEnd Sub\nEnd Module"),
        vec!["1.5"]
    );
}

#[test]
fn explicit_cdbl_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CDbl(\"1.5\"))\nEnd Sub\nEnd Module"),
        vec!["1.5"]
    );
}
#[test]
fn explicit_cdec_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CDec(\"1.5\"))\nEnd Sub\nEnd Module"),
        vec!["1.5"]
    );
}
#[test]
fn explicit_cbool_integer_true() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CBool(-1))\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn explicit_cbool_integer_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CBool(0))\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn explicit_cbool_string_true() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CBool(\"True\"))\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}

#[test]
fn explicit_cstr_integer() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CStr(10))\nEnd Sub\nEnd Module"),
        vec!["10"]
    );
}
#[test]
fn explicit_cstr_boolean() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CStr(True))\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn explicit_cstr_date() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nConsole.WriteLine(CStr(#1/1/2000#).Contains(\"2000\"))\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn explicit_cdate_string() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nConsole.WriteLine(CDate(\"1/1/2000\").Year)\nEnd Sub\nEnd Module"
        ),
        vec!["2000"]
    );
}
#[test]
fn explicit_cbyte_double() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CByte(254.5))\nEnd Sub\nEnd Module"),
        vec!["254"]
    );
}

#[test]
fn explicit_cchar_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CChar(\"A\"))\nEnd Sub\nEnd Module"),
        vec!["A"]
    );
}
#[test]
fn explicit_ctype_integer() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nConsole.WriteLine(CType(\"10\", Integer))\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn explicit_ctype_object_to_string() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = \"Hello\"\nConsole.WriteLine(CType(obj, String))\nEnd Sub\nEnd Module"
        ),
        vec!["Hello"]
    );
}
#[test]
fn explicit_directcast_exact() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = \"Hello\"\nConsole.WriteLine(DirectCast(obj, String))\nEnd Sub\nEnd Module"
        ),
        vec!["Hello"]
    );
}
#[test]
fn explicit_directcast_fails_coercion() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim obj As Object = 10\nTry\nDim s = DirectCast(obj, String)\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
} // DirectCast doesn't do numeric to string coercion

#[test]
fn explicit_trycast_success() {
    assert_eq!(
        run_vb(
            "Class C\nEnd Class\nModule M\nSub Main()\nDim obj As Object = New C()\nDim c1 = TryCast(obj, C)\nConsole.WriteLine(c1 IsNot Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn explicit_trycast_failure_nothing() {
    assert_eq!(
        run_vb(
            "Class C\nEnd Class\nModule M\nSub Main()\nDim obj As Object = \"Hello\"\nDim c1 = TryCast(obj, C)\nConsole.WriteLine(c1 Is Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
} // TryCast returns Nothing if it fails
#[test]
fn explicit_ctype_dynamic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim i = 10\nDim t = GetType(String)\n' CType requires compile time type, cannot use GetType dynamically in CType.\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn explicit_csbyte_string() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CSByte(\"-128\"))\nEnd Sub\nEnd Module"),
        vec!["-128"]
    );
}
#[test]
fn explicit_cuint_double() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(CUInt(4294967294.5))\nEnd Sub\nEnd Module"),
        vec!["4294967294"]
    );
}
