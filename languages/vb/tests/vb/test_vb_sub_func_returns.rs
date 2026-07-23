use super::helpers::run_vb;

#[test]
fn sub_call_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Print(v As String)\nConsole.WriteLine(v)\nEnd Sub\nSub Main()\nPrint(\"A\")\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn sub_call_call_keyword() {
    assert_eq!(
        run_vb(
            "Module M\nSub Print(v As String)\nConsole.WriteLine(v)\nEnd Sub\nSub Main()\nCall Print(\"B\")\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn sub_call_exit_sub() {
    assert_eq!(
        run_vb(
            "Module M\nSub Test(v As Integer)\nIf v = 1 Then Exit Sub\nConsole.WriteLine(v)\nEnd Sub\nSub Main()\nTest(1)\nTest(2)\nEnd Sub\nEnd Module"
        ),
        vec!["2"]
    );
}
#[test]
fn func_call_basic() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As String\nReturn \"C\"\nEnd Function\nSub Main()\nConsole.WriteLine(GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["C"]
    );
}
#[test]
fn func_call_implicit_return() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As Integer\nGetV = 42\nEnd Function\nSub Main()\nConsole.WriteLine(GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["42"]
    );
}

#[test]
fn func_call_implicit_return_exit() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As Integer\nGetV = 42\nExit Function\nEnd Function\nSub Main()\nConsole.WriteLine(GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["42"]
    );
}
#[test]
fn func_call_implicit_return_default() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As Integer\n' No assignment\nEnd Function\nSub Main()\nConsole.WriteLine(GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn func_call_call_keyword() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As Integer\nReturn 1\nEnd Function\nSub Main()\n' Using Call discards the result\nCall GetV()\nConsole.WriteLine(\"OK\")\nEnd Sub\nEnd Module"
        ),
        vec!["OK"]
    );
}
#[test]
fn func_call_return_object() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nFunction GetV() As Object\nReturn \"A\"\nEnd Function\nSub Main()\nConsole.WriteLine(GetV())\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn func_call_return_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As String\nReturn Nothing\nEnd Function\nSub Main()\nConsole.WriteLine(GetV() Is Nothing)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}

#[test]
fn func_call_as_statement_discard() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As Integer\nReturn 1\nEnd Function\nSub Main()\nGetV()\nConsole.WriteLine(\"OK\")\nEnd Sub\nEnd Module"
        ),
        vec!["OK"]
    );
}
#[test]
fn sub_signature_no_parens_call() {
    assert_eq!(
        run_vb(
            "Module M\nSub Print()\nConsole.WriteLine(\"A\")\nEnd Sub\nSub Main()\nPrint ' No parens\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn func_signature_no_parens_call() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetV() As String\nReturn \"A\"\nEnd Function\nSub Main()\nConsole.WriteLine(GetV)\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn sub_call_parentheses_single_arg() {
    assert_eq!(
        run_vb(
            "Module M\nSub Print(ByVal v As String)\nConsole.WriteLine(v)\nEnd Sub\nSub Main()\nPrint(\"A\")\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn sub_call_parentheses_single_arg_force_byval() {
    assert_eq!(
        run_vb(
            "Module M\nSub Mutate(ByRef v As Integer)\nv = 2\nEnd Sub\nSub Main()\nDim x = 1\nMutate((x)) ' Extra parens force evaluation as expression, bypassing ByRef\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
} // Classic VB behavior

#[test]
fn func_overload_resolution() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F(v As Integer) As String\nReturn \"Int\"\nEnd Function\nFunction F(v As String) As String\nReturn \"Str\"\nEnd Function\nSub Main()\nConsole.WriteLine(F(10))\nEnd Sub\nEnd Module"
        ),
        vec!["Int"]
    );
}
#[test]
fn func_overload_resolution_coercion() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nFunction F(v As Double) As String\nReturn \"Double\"\nEnd Function\nFunction F(v As String) As String\nReturn \"Str\"\nEnd Function\nSub Main()\nConsole.WriteLine(F(10)) ' Integer coerces to Double closer than String\nEnd Sub\nEnd Module"
        ),
        vec!["Double"]
    );
}
#[test]
fn func_recursive() {
    assert_eq!(
        run_vb(
            "Module M\nFunction Fact(n As Integer) As Integer\nIf n <= 1 Then Return 1\nReturn n * Fact(n - 1)\nEnd Function\nSub Main()\nConsole.WriteLine(Fact(3))\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn sub_recursive() {
    assert_eq!(
        run_vb(
            "Module M\nDim sum = 0\nSub Count(n As Integer)\nIf n <= 0 Then Exit Sub\nsum += 1\nCount(n - 1)\nEnd Sub\nSub Main()\nCount(3)\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn func_implicit_variable_mutation() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As Integer\nF = 10\nF += 5\nEnd Function\nSub Main()\nConsole.WriteLine(F())\nEnd Sub\nEnd Module"
        ),
        vec!["15"]
    );
} // The function name acts as a local variable

#[test]
fn func_return_array() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As Integer()\nReturn {1, 2}\nEnd Function\nSub Main()\nConsole.WriteLine(F().Length)\nEnd Sub\nEnd Module"
        ),
        vec!["2"]
    );
}
#[test]
fn func_return_nested_array() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As Integer(,)\nReturn {{1, 2}, {3, 4}}\nEnd Function\nSub Main()\nConsole.WriteLine(F()(1, 1))\nEnd Sub\nEnd Module"
        ),
        vec!["4"]
    );
}
#[test]
fn func_return_type_inference() {
    assert_eq!(
        run_vb(
            "Option Infer On\nModule M\nFunction F()\nReturn 10\nEnd Function\nSub Main()\nConsole.WriteLine(F().GetType().Name)\nEnd Sub\nEnd Module"
        ),
        vec!["Int32"]
    );
}
#[test]
fn func_shadowing_module_var() {
    assert_eq!(
        run_vb(
            "Module M\nDim F As Integer = 5\n' Cannot define a function with the same name as a module variable in the same module\nSub Main()\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn sub_return_statement_fails() {
    assert_eq!(
        run_vb(
            "Module M\nSub Print()\n' Return 10 ' Fails because Sub cannot return value\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nSub Main()\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
