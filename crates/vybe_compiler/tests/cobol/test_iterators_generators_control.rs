use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn iter_over_table_compiles() { compile_ok(&p("01 T PIC X(2) OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.", "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        DISPLAY T(I)\n    END-PERFORM.")); }
#[test] fn iter_search_pattern_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH E\n        WHEN K(I) = \"ABC\" DISPLAY \"F\"\n    END-SEARCH.")); }
#[test] fn iter_search_all_pattern_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH ALL E\n        WHEN K(I) = \"ABC\" DISPLAY \"F\"\n    END-SEARCH.")); }
#[test] fn iter_table_walk_with_accumulator_compiles() { compile_ok(&p("01 T PIC 9 OCCURS 3 TIMES.\n01 I PIC 9 VALUE 1.\n01 S PIC 99 VALUE 0.", "    MOVE 1 TO T(1).\n    MOVE 2 TO T(2).\n    MOVE 3 TO T(3).\n    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        ADD T(I) TO S\n    END-PERFORM.\n    DISPLAY S.")); }
#[test] fn iter_search_with_at_end_clause_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 3 TIMES INDEXED BY I.\n      10 K PIC X(3).", "    SEARCH E\n        AT END DISPLAY \"NONE\"\n        WHEN K(I) = \"ABC\" DISPLAY \"FOUND\"\n    END-SEARCH.")); }
#[test] fn iter_nested_table_traversal_compiles() { compile_ok(&p("01 I PIC 9 VALUE 1.\n01 J PIC 9 VALUE 1.\n01 T.\n   05 R OCCURS 2 TIMES.\n      10 C PIC 9 OCCURS 2 TIMES.", "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2\n        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2\n            MOVE J TO C(I J)\n        END-PERFORM\n    END-PERFORM.")); }
#[test] fn iter_search_all_numeric_key_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 4 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC 9(2).", "    SEARCH ALL E\n        WHEN K(I) = 10 DISPLAY \"HIT\"\n    END-SEARCH.")); }
#[test] fn iter_perform_until_with_table_condition_compiles() { compile_ok(&p("01 I PIC 9 VALUE 1.\n01 T PIC 9 OCCURS 3 TIMES.\n01 F PIC 9 VALUE 0.", "    MOVE 1 TO T(1). MOVE 2 TO T(2). MOVE 3 TO T(3).\n    PERFORM UNTIL I > 3 OR F = 1\n        IF T(I) = 2 MOVE 1 TO F END-IF\n        ADD 1 TO I\n    END-PERFORM.")); }
#[test] fn iter_varying_by_step_compiles() { compile_ok(&p("01 I PIC 9 VALUE 1.\n01 T PIC 9 OCCURS 5 TIMES.", "    PERFORM VARYING I FROM 1 BY 2 UNTIL I > 5\n        MOVE I TO T(I)\n    END-PERFORM.")); }
