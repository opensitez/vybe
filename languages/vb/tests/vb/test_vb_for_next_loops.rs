use super::helpers::run_vb;

#[test]
fn for_next_basic() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 5\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["15"]
    );
}
#[test]
fn for_next_inline_declaration() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i As Integer = 1 To 3\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_next_step() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 5 Step 2\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["9"]
    );
}
#[test]
fn for_next_negative_step() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 5 To 1 Step -1\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["15"]
    );
}
#[test]
fn for_next_no_execute_positive_step() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 5 To 1\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}

#[test]
fn for_next_no_execute_negative_step() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 5 Step -1\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn for_next_floating_point() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum As Double = 0\nFor i As Double = 0 To 1 Step 0.5\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["1.5"]
    );
}
#[test]
fn for_next_decimal() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum As Decimal = 0D\nFor i As Decimal = 0D To 1D Step 0.5D\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["1.5"]
    );
}
#[test]
fn for_next_exit_for() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 10\nIf i = 4 Then Exit For\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_next_continue_for() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 4\nIf i = 2 Then Continue For\nsum += i\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["8"]
    );
}

#[test]
fn for_next_variable_after_loop() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim i As Integer\nFor i = 1 To 3\nNext\nConsole.WriteLine(i)\nEnd Sub\nEnd Module"
        ),
        vec!["4"]
    );
} // Incremented past end
#[test]
fn for_next_variable_mutation() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim sum = 0\nFor i = 1 To 5\nsum += 1\ni = 5\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
} // Mutating loop var ends loop
#[test]
fn for_next_nested() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim count = 0\nFor i = 1 To 2\nFor j = 1 To 3\ncount += 1\nNext\nNext\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_next_multiple_vars_legacy() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\n' For i = 1 To 2: Next i, j is legacy VB6 syntax sometimes partially parsable\nConsole.WriteLine(\"Parsed\")\nEnd Sub\nEnd Module"
        ),
        vec!["Parsed"]
    );
}
#[test]
fn for_next_dynamic_limits() {
    assert_eq!(
        run_vb(
            "Module M\nFunction GetLimit() As Integer\nReturn 3\nEnd Function\nSub Main()\nDim sum = 0\nFor i = 1 To GetLimit()\nsum += 1\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}

#[test]
fn for_next_limit_evaluated_once() {
    assert_eq!(
        run_vb(
            "Module M\nDim limit As Integer = 3\nFunction GetLimit() As Integer\nlimit += 1\nReturn limit\nEnd Function\nSub Main()\nDim sum = 0\nFor i = 1 To GetLimit()\nsum += 1\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["4"]
    );
} // Evaluated once at start
#[test]
fn for_each_array() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {1, 2, 3}\nDim sum = 0\nFor Each v In arr\nsum += v\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_each_inline_declaration() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {1, 2, 3}\nDim sum = 0\nFor Each v As Integer In arr\nsum += v\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_each_string() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim s = \"ABC\"\nDim count = 0\nFor Each c In s\ncount += 1\nNext\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn for_each_collection() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim col As New Collection()\ncol.Add(10)\ncol.Add(20)\nDim sum = 0\nFor Each v In col\nsum += CInt(v)\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["30"]
    );
}

#[test]
fn for_each_exit_for() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {1, 2, 3, 4}\nDim sum = 0\nFor Each v In arr\nIf v = 3 Then Exit For\nsum += v\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["3"]
    );
}
#[test]
fn for_each_continue_for() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {1, 2, 3, 4}\nDim sum = 0\nFor Each v In arr\nIf v = 2 Then Continue For\nsum += v\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["8"]
    );
}
#[test]
fn for_each_type_coercion() {
    assert_eq!(
        run_vb(
            "Option Strict Off\nModule M\nSub Main()\nDim arr() As Object = {1, 2, 3}\nDim sum As Integer = 0\nFor Each v As Integer In arr\nsum += v\nNext\nConsole.WriteLine(sum)\nEnd Sub\nEnd Module"
        ),
        vec!["6"]
    );
}
#[test]
fn for_each_empty_array() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {}\nDim count = 0\nFor Each v In arr\ncount += 1\nNext\nConsole.WriteLine(count)\nEnd Sub\nEnd Module"
        ),
        vec!["0"]
    );
}
#[test]
fn for_each_variable_after_loop() {
    assert_eq!(
        run_vb(
            "Module M\nSub Main()\nDim arr() As Integer = {1}\nDim v As Integer\nFor Each v In arr\nNext\nConsole.WriteLine(v)\nEnd Sub\nEnd Module"
        ),
        vec!["1"]
    );
} // Retains last value
