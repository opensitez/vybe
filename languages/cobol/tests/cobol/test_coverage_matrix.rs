use super::helpers::{compile_ok, run_prints};

fn compile_case(index: usize, data: &str, body: &str) {
    let label = format!("CASE{:03}", index);
    let src = format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. {}.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        label, data, body
    );
    compile_ok(&src);
}

#[test]
fn cobol_numeric_arithmetic_matrix() {
    let pics = [
        "9(1)",
        "9(2)",
        "9(3)",
        "9(4)",
        "9(5)",
        "9(2)V9(2)",
        "S9(2)",
        "S9(3)",
        "9(3)V9(1)",
        "9(4)V9(2)",
    ];
    let values = [
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
        "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31",
        "32", "33", "34", "35", "36", "37", "38", "39", "40",
    ];
    let mut index = 0;
    for pic in pics {
        for left in values {
            for right in values {
                let data = format!(
                    "01 A PIC {} VALUE {}.\n01 B PIC {} VALUE {}.\n01 R PIC {} VALUE 0.",
                    pic, left, pic, right, pic
                );
                let body = format!(
                    "    ADD A TO B.\n    DISPLAY B.\n    SUBTRACT A FROM B.\n    DISPLAY B."
                );
                compile_case(index, &data, &body);
                index += 1;
            }
        }
    }
}

#[test]
fn cobol_alphanumeric_move_matrix() {
    let sizes = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 18, 20];
    let values = [
        "A", "B", "C", "AB", "CD", "HELLO", "WORLD", "COBOL", "123", "ABC123", "X1Y2", "MIXED",
        "VALUE", "LEFT", "RIGHT", "ALPHA", "BETA", "GAMMA", "DELTA", "EPSILON",
    ];
    let mut index = 0;
    for size in sizes {
        for value in values {
            let data = format!(
                "01 SRC PIC X({}) VALUE \"{}\".\n01 DST PIC X({}).\n01 ALT PIC X({}).",
                size, value, size, size
            );
            let body = format!(
                "    MOVE SRC TO DST.\n    DISPLAY DST.\n    STRING SRC DELIMITED BY SIZE INTO ALT.\n    DISPLAY ALT."
            );
            compile_case(index, &data, &body);
            index += 1;
        }
    }
}

#[test]
fn cobol_condition_matrix() {
    let lefts = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20, 25, 30];
    let rights = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20, 25, 30];
    let ops = ["=", "<", ">", "<=", ">="];
    let mut index = 0;
    for left in lefts {
        for right in rights {
            for op in ops {
                let data = format!(
                    "01 A PIC 9(2) VALUE {}.\n01 B PIC 9(2) VALUE {}.\n01 FLAG PIC X(1) VALUE \"N\".",
                    left, right
                );
                let body = format!(
                    "    IF A {} B MOVE \"Y\" TO FLAG END-IF.\n    DISPLAY FLAG.",
                    op
                );
                compile_case(index, &data, &body);
                index += 1;
            }
        }
    }
}

#[test]
fn cobol_evaluate_matrix() {
    let values = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
    let labels = [
        "ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
    ];
    let mut index = 0;
    for value in values {
        for label in labels {
            let data = format!(
                "01 N PIC 9(2) VALUE {}.\n01 OUT PIC X(10) VALUE \"NONE\".",
                value
            );
            let body = format!(
                "    EVALUATE N\n        WHEN {} DISPLAY \"{}\"\n        WHEN OTHER DISPLAY \"OTHER\"\n    END-EVALUATE.",
                value, label
            );
            compile_case(index, &data, &body);
            index += 1;
        }
    }
}

#[test]
fn cobol_inspect_string_matrix() {
    let texts = [
        "ALPHA", "BETA", "GAMMA", "DELTA", "EPSILON", "ZETA", "ETA", "THETA", "IOTA", "KAPPA",
    ];
    let tokens = ["A", "B", "C", "D", "E", "G", "H", "L", "M", "T"];
    let modes = ["ALL", "LEADING", "FIRST"];
    let mut index = 0;
    for text in texts {
        for token in tokens {
            for mode in modes {
                let data = format!(
                    "01 TXT PIC X(20) VALUE \"{}\".\n01 CNT PIC 9(3) VALUE 0.",
                    text
                );
                let body = format!(
                    "    INSPECT TXT TALLYING CNT FOR {} \"{}\".\n    DISPLAY CNT.",
                    mode, token
                );
                compile_case(index, &data, &body);
                index += 1;
            }
        }
    }
}

#[test]
fn cobol_reference_modification_matrix() {
    let texts = [
        "HELLO", "WORLD", "COBOL", "PROGRAM", "ROUTINE", "LITERAL", "MATRIX", "COVERAGE",
        "COMPILE", "TARGET",
    ];
    let starts = [1, 2, 3, 4, 5, 6, 7];
    let lengths = [1, 2, 3, 4, 5, 6, 7, 8];
    let mut index = 0;
    for text in texts {
        for start in starts {
            for length in lengths {
                let data = format!("01 SRC PIC X(30) VALUE \"{}\".\n01 SUB PIC X(10).", text);
                let body = format!(
                    "    MOVE SRC({}:{}) TO SUB.\n    DISPLAY SUB.",
                    start, length
                );
                compile_case(index, &data, &body);
                index += 1;
            }
        }
    }
}

#[test]
fn cobol_redefines_and_level_matrix() {
    let values = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"];
    let mut index = 0;
    for value in values {
        let data = format!(
            "01 MAIN PIC X(10) VALUE \"{}\".\n01 ALT REDEFINES MAIN PIC 9(2).\n01 NUM PIC 9(2) VALUE 0.",
            value
        );
        let body = format!("    MOVE ALT TO NUM.\n    DISPLAY NUM.");
        compile_case(index, &data, &body);
        index += 1;
    }
}

#[test]
fn cobol_figure_and_initialize_matrix() {
    let values = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
    let mut index = 0;
    for value in values {
        let data = format!(
            "01 FILL PIC X(10) VALUE \"{}\".\n01 CLEAN PIC X(10) VALUE \"ZZZ\".",
            value
        );
        let body = format!("    INITIALIZE CLEAN.\n    DISPLAY CLEAN.");
        compile_case(index, &data, &body);
        index += 1;
    }
}

#[test]
fn numeric_arithmetic_single_case_runtime() {
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. SAMPLE.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(2) VALUE 2.\n01 B PIC 9(2) VALUE 3.\n01 R PIC 9(3) VALUE 0.\nPROCEDURE DIVISION.\n    ADD A TO B\n    DISPLAY B\n    SUBTRACT A FROM B\n    DISPLAY B\n    STOP RUN.";
    let out = run_prints(src);
    assert_eq!(out, vec!["5", "3"]);
}

#[test]
fn condition_matrix_single_runtime() {
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. SAMPLE2.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9(2) VALUE 7.\n01 B PIC 9(2) VALUE 7.\n01 FLAG PIC X(1) VALUE \"N\".\nPROCEDURE DIVISION.\n    IF A = B MOVE \"Y\" TO FLAG END-IF.\n    DISPLAY FLAG\n    STOP RUN.";
    let out = run_prints(src);
    assert_eq!(out, vec!["Y"]);
}
