use super::helpers::run_vb;

#[test]
fn log_and_true_true() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(True And True)\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn log_and_true_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(True And False)\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn log_or_true_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(True Or False)\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn log_or_false_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(False Or False)\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn log_xor_true_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(True Xor False)\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}

#[test]
fn log_xor_true_true() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(True Xor True)\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn log_not_true() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(Not True)\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
}
#[test]
fn log_not_false() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(Not False)\nEnd Sub\nEnd Module"),
        vec!["True"]
    );
}
#[test]
fn bitwise_and_integers() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(12 And 10)\nEnd Sub\nEnd Module"),
        vec!["8"]
    );
}
#[test]
fn bitwise_or_integers() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(12 Or 10)\nEnd Sub\nEnd Module"),
        vec!["14"]
    );
}

#[test]
fn bitwise_xor_integers() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(12 Xor 10)\nEnd Sub\nEnd Module"),
        vec!["6"]
    );
}
#[test]
fn bitwise_not_integer() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(Not 0)\nEnd Sub\nEnd Module"),
        vec!["-1"]
    );
} // Two's complement not
#[test]
fn log_andalso_short_circuit() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = False AndAlso (1 / 0 > 0)\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["False"]
    );
}
#[test]
fn log_orelse_short_circuit() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = True OrElse (1 / 0 > 0)\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
#[test]
fn log_and_no_short_circuit_throws() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nDim x = False And (1 \\ 0 > 0)\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}

#[test]
fn log_or_no_short_circuit_throws() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nTry\nDim x = True Or (1 \\ 0 > 0)\nCatch\nConsole.WriteLine(\"Caught\")\nEnd Try\nEnd Sub\nEnd Module"
        ),
        vec!["Caught"]
    );
}
#[test]
fn log_andalso_evaluation_order() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As Boolean\nConsole.WriteLine(\"F\")\nReturn False\nEnd Function\nSub Main()\nDim x = F() AndAlso F()\nEnd Sub\nEnd Module"
        ),
        vec!["F"]
    );
} // Only evaluated once
#[test]
fn log_orelse_evaluation_order() {
    assert_eq!(
        run_vb(
            "Module M\nFunction F() As Boolean\nConsole.WriteLine(\"F\")\nReturn True\nEnd Function\nSub Main()\nDim x = F() OrElse F()\nEnd Sub\nEnd Module"
        ),
        vec!["F"]
    );
} // Only evaluated once
#[test]
fn bitwise_shift_left() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(1 << 2)\nEnd Sub\nEnd Module"),
        vec!["4"]
    );
}
#[test]
fn bitwise_shift_right() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(8 >> 2)\nEnd Sub\nEnd Module"),
        vec!["2"]
    );
}

#[test]
fn bitwise_shift_left_assignment() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 1\nx <<= 3\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["8"]
    );
}
#[test]
fn bitwise_shift_right_assignment() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim x = 16\nx >>= 2\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["4"]
    );
}
#[test]
fn log_precedence_and_or() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nConsole.WriteLine(True Or False And False)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
} // And has higher precedence than Or
#[test]
fn log_precedence_not_and() {
    assert_eq!(
        run_vb("Module M\nSub Main()\nConsole.WriteLine(Not False And False)\nEnd Sub\nEnd Module"),
        vec!["False"]
    );
} // Not has higher precedence than And
#[test]
fn log_implicit_boolean_conversion() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim x = \"True\" And \"True\"\nConsole.WriteLine(x)\nEnd Sub\nEnd Module"
        ),
        vec!["True"]
    );
}
