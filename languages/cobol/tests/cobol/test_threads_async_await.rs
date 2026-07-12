use super::helpers::compile_ok;

#[test] fn call_statement_basic_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUBA\".\n    STOP RUN."); }
#[test] fn call_statement_using_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 V PIC 9(3) VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"SUBB\" USING V.\n    STOP RUN."); }
#[test] fn call_with_on_exception_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUBC\"\n        ON EXCEPTION DISPLAY \"E\"\n        NOT ON EXCEPTION DISPLAY \"O\"\n    END-CALL.\n    STOP RUN."); }
#[test] fn perform_times_with_call_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM 2 TIMES\n        CALL \"SUBD\"\n    END-PERFORM.\n    STOP RUN."); }
#[test] fn perform_until_with_call_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 I PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM UNTIL I >= 2\n        ADD 1 TO I\n        CALL \"SUBE\"\n    END-PERFORM.\n    STOP RUN."); }
#[test] fn evaluate_with_call_branches_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 K PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    EVALUATE K\n        WHEN 1 CALL \"SUB1\"\n        WHEN 2 CALL \"SUB2\"\n        WHEN OTHER CALL \"SUBX\"\n    END-EVALUATE.\n    STOP RUN."); }
#[test] fn if_else_with_call_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 F PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    IF F = 1\n        CALL \"SUBY\"\n    ELSE\n        CALL \"SUBZ\"\n    END-IF.\n    STOP RUN."); }
#[test] fn call_chain_two_programs_compiles() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUBM\".\n    CALL \"SUBN\".\n    STOP RUN."); }
