use vybe_parser::parse_program;
use vybe_runtime::Interpreter;

fn test(name: &str, code: &str) {
    print!("Testing {}... ", name);
    let program = match parse_program(code) {
        Ok(p) => p,
        Err(e) => { println!("PARSE ERROR: {}", e); return; }
    };
    let mut interp = Interpreter::new();
    match interp.run(&program) {
        Ok(()) => {
            let mut results = Vec::new();
            while let Some(effect) = interp.side_effects.pop_front() {
                match effect {
                    vybe_runtime::RuntimeSideEffect::MsgBox(msg) => results.push(msg),
                    _ => {}
                }
            }
            println!("OK -> {}", results.join(", "));
        },
        Err(e) => println!("RUNTIME ERROR: {}", e),
    }
}

fn main() {
    println!("=== vybe Feature Tests ===\n");

    test("Basic", "Dim x As Integer\nx = 42\nMsgBox(x)\n");

    test("Enum", "Enum Colors\n    Red\n    Green = 5\n    Blue\nEnd Enum\nMsgBox(\"Red=\" & Red & \" Green=\" & Green & \" Blue=\" & Blue)\n");

    test("ForEach/Array", "Dim arr() As Variant\narr = Array(10, 20, 30, 40, 50)\nDim total As Integer\ntotal = 0\nDim item As Variant\nFor Each item In arr\n    total = total + item\nNext\nMsgBox(\"total=\" & total)\n");

    test("ForEach/String", "Dim s As String\ns = \"ABC\"\nDim result As String\nresult = \"\"\nDim ch As String\nFor Each ch In s\n    result = result & ch & \"-\"\nNext\nMsgBox(result)\n");

    test("With", "Class Box\n    Public Width As Integer\n    Public Height As Integer\nEnd Class\nDim b As Box\nSet b = New Box\nWith b\n    .Width = 100\n    .Height = 200\nEnd With\nMsgBox(\"W=\" & b.Width & \" H=\" & b.Height)\n");

    test("StringFns", "MsgBox(Trim(\"  hello  \") & \"|\" & Len(\"abc\") & \"|\" & Mid(\"hello\", 2, 3))\n");

    test("MathFns", "MsgBox(Abs(-5) & \"|\" & Sqr(16) & \"|\" & Round(3.14159, 2))\n");

    test("SelectCase", "Dim x As Integer\nx = 2\nSelect Case x\n    Case 1\n        MsgBox(\"one\")\n    Case 2, 3\n        MsgBox(\"two or three\")\n    Case Else\n        MsgBox(\"other\")\nEnd Select\n");

    test("TryCatch", "Try\n    Dim x As Integer\n    x = CInt(\"bad\")\nCatch ex As Exception\n    MsgBox(\"caught\")\nEnd Try\n");

    test("NewMath", "MsgBox(Max(10, 20) & \"|\" & Min(10, 20) & \"|\" & Pow(2, 10) & \"|\" & Ceiling(3.2) & \"|\" & Floor(3.8))\n");

    test("NewConv", "MsgBox(CByte(42) & \"|\" & FormatNumber(1234.5, 1) & \"|\" & FormatPercent(0.75, 0))\n");

    test("RGB", "MsgBox(RGB(255, 0, 0))\n");

    test("IsNullOrEmpty", "MsgBox(IsNullOrEmpty(\"\") & \"|\" & IsNullOrEmpty(\"hi\"))\n");

    test("While/Wend", "Dim n As Integer\nn = 0\nWhile n < 3\n    n = n + 1\nWend\nMsgBox(\"n=\" & n)\n");

    println!("\n=== Done ===");
}
