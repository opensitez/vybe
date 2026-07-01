use super::helpers::compile_ok;

fn p(body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT F ASSIGN TO \"f.dat\" ORGANIZATION IS SEQUENTIAL.\nDATA DIVISION.\nFILE SECTION.\nFD F.\n01 R PIC X(20).\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        body
    )
}

#[test]
fn sequential_open_write_close_compiles() {
    compile_ok(&p("    OPEN OUTPUT F.\n    WRITE R.\n    CLOSE F."));
}

#[test]
fn sequential_open_read_at_end_compiles() {
    compile_ok(&p("    OPEN INPUT F.\n    READ F AT END DISPLAY \"EOF\" END-READ.\n    CLOSE F."));
}

#[test]
fn sequential_open_input_compiles() {
    compile_ok(&p("    OPEN INPUT F.\n    CLOSE F."));
}

#[test]
fn sequential_open_io_compiles() {
    compile_ok(&p("    OPEN I-O F.\n    CLOSE F."));
}

#[test]
fn sequential_open_extend_compiles() {
    compile_ok(&p("    OPEN EXTEND F.\n    CLOSE F."));
}

#[test]
fn sequential_read_into_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT F ASSIGN TO \"f.dat\" ORGANIZATION IS SEQUENTIAL.\nDATA DIVISION.\nFILE SECTION.\nFD F.\n01 R PIC X(20).\nWORKING-STORAGE SECTION.\n01 H PIC X(20).\nPROCEDURE DIVISION.\n    OPEN INPUT F.\n    READ F INTO H AT END DISPLAY \"EOF\" END-READ.\n    CLOSE F.\n    STOP RUN.");
}

#[test]
fn sequential_write_from_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nENVIRONMENT DIVISION.\nINPUT-OUTPUT SECTION.\nFILE-CONTROL.\n    SELECT F ASSIGN TO \"f.dat\" ORGANIZATION IS SEQUENTIAL.\nDATA DIVISION.\nFILE SECTION.\nFD F.\n01 R PIC X(20).\nWORKING-STORAGE SECTION.\n01 H PIC X(20).\nPROCEDURE DIVISION.\n    OPEN OUTPUT F.\n    WRITE R FROM H.\n    CLOSE F.\n    STOP RUN.");
}

#[test]
fn sequential_write_advancing_lines_compiles() {
    compile_ok(&p("    OPEN OUTPUT F.\n    WRITE R AFTER ADVANCING 2 LINES.\n    CLOSE F."));
}

#[test]
fn sequential_write_advancing_page_compiles() {
    compile_ok(&p("    OPEN OUTPUT F.\n    WRITE R AFTER ADVANCING PAGE.\n    CLOSE F."));
}

#[test]
fn sequential_close_with_no_rewind_compiles() {
    compile_ok(&p("    OPEN INPUT F.\n    CLOSE F WITH NO REWIND."));
}