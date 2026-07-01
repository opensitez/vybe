use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn add_case_01() { compile_ok(&p("01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.", "    ADD A TO B.")); }
#[test] fn add_case_02() { compile_ok(&p("01 A PIC 9 VALUE 3.\n01 B PIC 9 VALUE 4.", "    ADD 2 TO B.")); }
#[test] fn add_case_03() { compile_ok(&p("01 A PIC 9 VALUE 1.\n01 B PIC 9 VALUE 2.\n01 C PIC 9 VALUE 3.", "    ADD A B TO C.")); }
#[test] fn add_case_04() { compile_ok(&p("01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 1.\n01 R PIC 9.", "    ADD A B GIVING R.")); }
#[test] fn add_case_05() { compile_ok(&p("01 A PIC 9 VALUE 7.\n01 B PIC 9 VALUE 1.", "    ADD B TO A END-ADD.")); }
#[test] fn add_case_06() { compile_ok(&p("01 A PIC 9 VALUE 9.\n01 R1 PIC 99.\n01 R2 PIC 99.", "    ADD A 1 GIVING R1 R2.")); }
#[test] fn add_case_07() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 2.", "    ADD A TO B ON SIZE ERROR DISPLAY \"E\" END-ADD.")); }
#[test] fn add_case_08() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 2.", "    ADD A TO B NOT ON SIZE ERROR DISPLAY \"O\" END-ADD.")); }
#[test] fn add_case_09() { compile_ok(&p("01 A PIC 9V9 VALUE 1.5.\n01 B PIC 9V9 VALUE 2.5.\n01 R PIC 9.", "    ADD A B GIVING R ROUNDED.")); }
#[test] fn add_case_10() { compile_ok(&p("01 G1.\n   05 A PIC 9 VALUE 1.\n01 G2.\n   05 A PIC 9 VALUE 2.", "    ADD CORRESPONDING G1 TO G2.")); }

#[test] fn sub_case_01() { compile_ok(&p("01 A PIC 9 VALUE 8.\n01 B PIC 9 VALUE 3.", "    SUBTRACT B FROM A.")); }
#[test] fn sub_case_02() { compile_ok(&p("01 A PIC 99 VALUE 20.", "    SUBTRACT 5 FROM A.")); }
#[test] fn sub_case_03() { compile_ok(&p("01 A PIC 99 VALUE 9.\n01 B PIC 9 VALUE 3.\n01 R PIC 99.", "    SUBTRACT B FROM A GIVING R.")); }
#[test] fn sub_case_04() { compile_ok(&p("01 A PIC 99 VALUE 9.\n01 B PIC 9 VALUE 2.", "    SUBTRACT B FROM A END-SUBTRACT.")); }
#[test] fn sub_case_05() { compile_ok(&p("01 A PIC 9V9 VALUE 3.5.\n01 B PIC 9V9 VALUE 1.2.\n01 R PIC 9.", "    SUBTRACT B FROM A GIVING R ROUNDED.")); }
#[test] fn sub_case_06() { compile_ok(&p("01 A PIC 9 VALUE 7.\n01 B PIC 9 VALUE 1.", "    SUBTRACT B FROM A ON SIZE ERROR DISPLAY \"E\" END-SUBTRACT.")); }
#[test] fn sub_case_07() { compile_ok(&p("01 A PIC 9 VALUE 7.\n01 B PIC 9 VALUE 1.", "    SUBTRACT B FROM A NOT ON SIZE ERROR DISPLAY \"O\" END-SUBTRACT.")); }
#[test] fn sub_case_08() { compile_ok(&p("01 G1.\n   05 A PIC 9 VALUE 8.\n01 G2.\n   05 A PIC 9 VALUE 2.", "    SUBTRACT CORRESPONDING G2 FROM G1.")); }
#[test] fn sub_case_09() { compile_ok(&p("01 A PIC S9 VALUE -2.\n01 B PIC 9 VALUE 1.", "    SUBTRACT B FROM A.")); }
#[test] fn sub_case_10() { compile_ok(&p("01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 0.", "    SUBTRACT B FROM A.")); }

#[test] fn mul_case_01() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.", "    MULTIPLY A BY B.")); }
#[test] fn mul_case_02() { compile_ok(&p("01 A PIC 9 VALUE 4.", "    MULTIPLY 2 BY A.")); }
#[test] fn mul_case_03() { compile_ok(&p("01 A PIC 9 VALUE 6.\n01 B PIC 9 VALUE 7.\n01 R PIC 99.", "    MULTIPLY A BY B GIVING R.")); }
#[test] fn mul_case_04() { compile_ok(&p("01 A PIC 9 VALUE 3.\n01 B PIC 9 VALUE 2.", "    MULTIPLY A BY B END-MULTIPLY.")); }
#[test] fn mul_case_05() { compile_ok(&p("01 A PIC 9V9 VALUE 1.5.\n01 B PIC 9V9 VALUE 2.5.\n01 R PIC 9.", "    MULTIPLY A BY B GIVING R ROUNDED.")); }
#[test] fn mul_case_06() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 2.", "    MULTIPLY A BY B ON SIZE ERROR DISPLAY \"E\" END-MULTIPLY.")); }
#[test] fn mul_case_07() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 2.", "    MULTIPLY A BY B NOT ON SIZE ERROR DISPLAY \"O\" END-MULTIPLY.")); }
#[test] fn mul_case_08() { compile_ok(&p("01 A PIC 9 VALUE 0.\n01 B PIC 9 VALUE 9.", "    MULTIPLY A BY B.")); }

#[test] fn div_case_01() { compile_ok(&p("01 A PIC 99 VALUE 20.\n01 B PIC 9 VALUE 5.", "    DIVIDE B INTO A.")); }
#[test] fn div_case_02() { compile_ok(&p("01 A PIC 99 VALUE 20.", "    DIVIDE 2 INTO A.")); }
#[test] fn div_case_03() { compile_ok(&p("01 A PIC 99 VALUE 20.\n01 B PIC 9 VALUE 5.\n01 R PIC 99.", "    DIVIDE A BY B GIVING R.")); }
#[test] fn div_case_04() { compile_ok(&p("01 A PIC 99 VALUE 20.\n01 B PIC 9 VALUE 3.\n01 Q PIC 99.\n01 M PIC 9.", "    DIVIDE B INTO A GIVING Q REMAINDER M.")); }
#[test] fn div_case_05() { compile_ok(&p("01 A PIC 99 VALUE 12.\n01 B PIC 9 VALUE 3.", "    DIVIDE B INTO A END-DIVIDE.")); }
#[test] fn div_case_06() { compile_ok(&p("01 A PIC 99 VALUE 12.\n01 B PIC 9 VALUE 3.", "    DIVIDE B INTO A ON SIZE ERROR DISPLAY \"E\" END-DIVIDE.")); }
#[test] fn div_case_07() { compile_ok(&p("01 A PIC 99 VALUE 12.\n01 B PIC 9 VALUE 3.", "    DIVIDE B INTO A NOT ON SIZE ERROR DISPLAY \"O\" END-DIVIDE.")); }
#[test] fn div_case_08() { compile_ok(&p("01 A PIC 99 VALUE 18.\n01 B PIC 9 VALUE 6.", "    DIVIDE A BY B GIVING A.")); }

#[test] fn cmp_case_01() { compile_ok(&p("01 R PIC 999 VALUE 0.", "    COMPUTE R = 1 + 2 + 3.")); }
#[test] fn cmp_case_02() { compile_ok(&p("01 R PIC 999 VALUE 0.", "    COMPUTE R = (2 + 3) * 4.")); }
#[test] fn cmp_case_03() { compile_ok(&p("01 R PIC S999 VALUE 0.", "    COMPUTE R = -5 + 1.")); }
#[test] fn cmp_case_04() { compile_ok(&p("01 R PIC 999 VALUE 0.", "    COMPUTE R = 20 / 4.")); }
#[test] fn cmp_case_05() { compile_ok(&p("01 R PIC 999 VALUE 0.", "    COMPUTE R = 2 * (3 + 4).")); }
#[test] fn cmp_case_06() { compile_ok(&p("01 R PIC 999 VALUE 0.", "    COMPUTE R = 9 - 7 + 5.")); }
#[test] fn cmp_case_07() { compile_ok(&p("01 R PIC 9 VALUE 0.", "    COMPUTE R = 1 + 1 ON SIZE ERROR DISPLAY \"E\" END-COMPUTE.")); }
#[test] fn cmp_case_08() { compile_ok(&p("01 R PIC 9 VALUE 0.", "    COMPUTE R = 1 + 1 NOT ON SIZE ERROR DISPLAY \"O\" END-COMPUTE.")); }
#[test] fn cmp_case_09() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.\n01 R PIC 99 VALUE 0.", "    COMPUTE R = A * B + 1.")); }
#[test] fn cmp_case_10() { compile_ok(&p("01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.\n01 C PIC 9 VALUE 4.\n01 R PIC 999 VALUE 0.", "    COMPUTE R = (A + B) * C.")); }