use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn keyed_table_decl_compiles() { compile_ok(&p("01 MAP.\n   05 ENTRY OCCURS 5 TIMES ASCENDING KEY IS K INDEXED BY I.\n      10 K PIC X(5).\n      10 V PIC X(10).", "    MOVE \"A\" TO K(1).")); }
#[test] fn keyed_insert_pattern_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"A\".\n01 V PIC X(10) VALUE \"ONE\".", "    CALL \"MAP-PUT\" USING K V.")); }
#[test] fn keyed_lookup_pattern_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"A\".\n01 V PIC X(10).", "    CALL \"MAP-GET\" USING K V.")); }
#[test] fn keyed_delete_pattern_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"A\".", "    CALL \"MAP-DEL\" USING K.")); }
#[test] fn keyed_exists_pattern_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"A\".\n01 E PIC 9 VALUE 0.", "    CALL \"MAP-EXISTS\" USING K E.")); }
#[test] fn keyed_size_pattern_compiles() { compile_ok(&p("01 N PIC 9(5).", "    CALL \"MAP-SIZE\" USING N.")); }
#[test] fn keyed_clear_pattern_compiles() { compile_ok(&p("", "    CALL \"MAP-CLEAR\".")); }
#[test] fn keyed_iterate_pattern_compiles() { compile_ok(&p("01 I PIC 9 VALUE 1.", "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5\n        CALL \"MAP-NEXT\" USING I\n    END-PERFORM.")); }
#[test] fn map_search_by_key_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES INDEXED BY I.\n      10 K PIC X(5).", "    SEARCH E\n        WHEN K(I) = \"A\" DISPLAY \"F\"\n    END-SEARCH.")); }
#[test] fn map_search_all_by_key_compiles() { compile_ok(&p("01 T.\n   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.\n      10 K PIC X(5).", "    SEARCH ALL E\n        WHEN K(I) = \"A\" DISPLAY \"F\"\n    END-SEARCH.")); }
#[test] fn map_value_update_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"A\".\n01 V PIC X(10) VALUE \"TWO\".", "    CALL \"MAP-PUT\" USING K V.")); }
#[test] fn map_put_if_absent_compiles() { compile_ok(&p("01 K PIC X(5) VALUE \"B\".\n01 V PIC X(10) VALUE \"X\".", "    CALL \"MAP-PUT-IF-ABSENT\" USING K V.")); }
#[test] fn map_merge_pattern_compiles() { compile_ok(&p("", "    CALL \"MAP-MERGE\".")); }
#[test] fn map_keys_iteration_compiles() { compile_ok(&p("", "    CALL \"MAP-KEYS\".")); }
#[test] fn map_values_iteration_compiles() { compile_ok(&p("", "    CALL \"MAP-VALUES\".")); }
#[test] fn map_entries_iteration_compiles() { compile_ok(&p("", "    CALL \"MAP-ENTRIES\".")); }
#[test] fn map_from_json_pattern_compiles() { compile_ok(&p("01 J PIC X(100).", "    CALL \"MAP-FROM-JSON\" USING J.")); }
#[test] fn map_to_json_pattern_compiles() { compile_ok(&p("01 J PIC X(100).", "    CALL \"MAP-TO-JSON\" USING J.")); }
