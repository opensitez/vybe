use super::helpers::run_vb;

#[test]
fn select_case_basic_match() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 2\nSelect Case x\nCase 1\nConsole.WriteLine(\"A\")\nCase 2\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn select_case_else() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 5\nSelect Case x\nCase 1\nConsole.WriteLine(\"A\")\nCase Else\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn select_case_multiple_values() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 3\nSelect Case x\nCase 1, 2, 3\nConsole.WriteLine(\"A\")\nCase Else\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_to_range() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 15\nSelect Case x\nCase 1 To 10\nConsole.WriteLine(\"A\")\nCase 11 To 20\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn select_case_is_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 25\nSelect Case x\nCase Is > 20\nConsole.WriteLine(\"A\")\nCase Else\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}

#[test]
fn select_case_string_match() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"Hello\"\nSelect Case s\nCase \"Hi\"\nConsole.WriteLine(\"A\")\nCase \"Hello\"\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn select_case_string_is_operator() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"Z\"\nSelect Case s\nCase Is > \"M\"\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_mixed_conditions() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 5\nSelect Case x\nCase 1, 3 To 7, Is > 10\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_boolean_expression() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 5\nSelect Case True\nCase x = 5\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_evaluation_order() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F(v As Integer) As Integer\nConsole.WriteLine(v)\nReturn v\nEnd Function\nSub Main()\nSelect Case F(5)\nCase F(1), F(5), F(10)\nConsole.WriteLine(\"Match\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["5", "1", "5", "Match"]
    );
}

#[test]
fn select_case_fallthrough_not_allowed() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1\nSelect Case x\nCase 1\nConsole.WriteLine(\"A\")\n' No fallthrough to Case 2 implicitly\nCase 2\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_empty() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1\nSelect Case x\nEnd Select\nConsole.WriteLine(\"OK\")\nEnd Sub\nEnd Module"
        ),
        vec!["OK"]
    );
}
#[test]
fn select_case_nested() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1, y = 2\nSelect Case x\nCase 1\nSelect Case y\nCase 2\nConsole.WriteLine(\"Nested\")\nEnd Select\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["Nested"]
    );
}
#[test]
fn select_case_type_coercion() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim x As String = \"5\"\nSelect Case x\nCase 1 To 10\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
} // "5" implicitly coerced to Integer for the 1 To 10 range check
#[test]
fn select_case_nothing() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x As Object = Nothing\nSelect Case x\nCase Nothing\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}

#[test]
fn select_case_exit_select() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1\nSelect Case x\nCase 1\nConsole.WriteLine(\"A\")\nExit Select\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_variable_declaration() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nSelect Case 1\nCase 1\nDim y = 10\nConsole.WriteLine(y)\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["10"]
    );
}
#[test]
fn select_case_enum() {
    assert_eq!(
        run_vb(
            "Enum E\nA\nB\nEnd Enum\nModule M\nSub Main()\nDim x = E.B\nSelect Case x\nCase E.A\nConsole.WriteLine(\"A\")\nCase E.B\nConsole.WriteLine(\"B\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["B"]
    );
}
#[test]
fn select_case_constant_expression() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nConst v = 5\nSelect Case 5\nCase v\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_multiple_is() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = -5\nSelect Case x\nCase Is < 0, Is > 10\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}

#[test]
fn select_case_type_mismatch_throws() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim x = \"ABC\"\nTry\nSelect Case x\nCase 1 To 10\nConsole.WriteLine(\"A\")\nEnd Select\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn select_case_boolean_false() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nSelect Case False\nCase True\nConsole.WriteLine(\"T\")\nCase False\nConsole.WriteLine(\"F\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["F"]
    );
}
#[test]
fn select_case_single_line_statements() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nSelect Case 1: Case 1: Console.WriteLine(\"A\"): Case 2: Console.WriteLine(\"B\"): End Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
#[test]
fn select_case_is_equal_implicit() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nSelect Case 5\nCase Is = 5\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
} // "Is =" is valid syntax, equivalent to "Case 5"
#[test]
fn select_case_is_not_equal() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nSelect Case 5\nCase Is <> 4\nConsole.WriteLine(\"A\")\nEnd Select\nEnd Sub\nEnd Module"
        ),
        vec!["A"]
    );
}
