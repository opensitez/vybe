use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT WS-FILE ASSIGN TO \"tmp.dat\".\nDATA DIVISION.\nFILE SECTION.\nFD WS-FILE.\n01 WS-REC PIC X(80).\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn open_input_compiles() { compile_ok(&p("", "    OPEN INPUT WS-FILE.")); }
#[test] fn open_output_compiles() { compile_ok(&p("", "    OPEN OUTPUT WS-FILE.")); }
#[test] fn open_io_compiles() { compile_ok(&p("", "    OPEN I-O WS-FILE.")); }
#[test] fn open_extend_compiles() { compile_ok(&p("", "    OPEN EXTEND WS-FILE.")); }
#[test] fn close_file_compiles() { compile_ok(&p("", "    CLOSE WS-FILE.")); }
#[test] fn read_file_compiles() { compile_ok(&p("", "    READ WS-FILE INTO WS-REC.")); }
#[test] fn write_file_compiles() { compile_ok(&p("", "    WRITE WS-REC.")); }
#[test] fn rewrite_file_compiles() { compile_ok(&p("", "    REWRITE WS-REC.")); }
#[test] fn delete_file_record_compiles() { compile_ok(&p("", "    DELETE WS-FILE.")); }
#[test] fn start_file_key_compiles() { compile_ok(&p("", "    START WS-FILE KEY IS = WS-REC.")); }
#[test] fn read_next_compiles() { compile_ok(&p("", "    READ WS-FILE NEXT RECORD INTO WS-REC.")); }
#[test] fn read_with_at_end_compiles() { compile_ok(&p("", "    READ WS-FILE\n        AT END DISPLAY \"EOF\"\n    END-READ.")); }
#[test] fn write_with_invalid_key_compiles() { compile_ok(&p("", "    WRITE WS-REC\n        INVALID KEY DISPLAY \"ERR\"\n    END-WRITE.")); }
#[test] fn open_read_close_sequence_compiles() { compile_ok(&p("", "    OPEN INPUT WS-FILE.\n    READ WS-FILE INTO WS-REC.\n    CLOSE WS-FILE.")); }
#[test] fn open_write_close_sequence_compiles() { compile_ok(&p("", "    OPEN OUTPUT WS-FILE.\n    WRITE WS-REC.\n    CLOSE WS-FILE.")); }
#[test] fn open_rewrite_close_sequence_compiles() { compile_ok(&p("", "    OPEN I-O WS-FILE.\n    REWRITE WS-REC.\n    CLOSE WS-FILE.")); }
#[test] fn sort_file_on_key_compiles() { compile_ok(&p("01 WS-KEY PIC 9(5).", "    SORT WS-FILE ON ASCENDING KEY WS-KEY.")); }
#[test] fn merge_file_on_key_compiles() { compile_ok(&p("01 WS-KEY PIC 9(5).", "    MERGE WS-FILE ON DESCENDING KEY WS-KEY.")); }
