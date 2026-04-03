/// Debug: check if COBOL data items parse correctly in different contexts

#[test]
fn parse_data_items_standalone() {
    // Minimal COBOL — check if data items are parsed
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-X PIC X(10) VALUE \"Hello\".\nPROCEDURE DIVISION.\n    STOP RUN.";

    // First verify tokenization
    let mut lexer = vybe_parser_cobol::lexer::Lexer::new(src);
    let tokens = lexer.tokenize().expect("tokenize failed");
    let token_kinds: Vec<String> = tokens.iter().map(|t| format!("{:?}", t.kind)).collect();

    // Find where DATA token is
    let data_pos = tokens.iter().position(|t| t.kind == vybe_parser_cobol::token::TokenKind::Data);
    let ws_pos = tokens.iter().position(|t| t.kind == vybe_parser_cobol::token::TokenKind::WorkingStorage);

    let prog = vybe_parser_cobol::parse(src).expect("parse failed");
    assert!(!prog.data_items.is_empty(),
        "Should have data items, got {}. DATA at {:?}, WS at {:?}. Tokens: {:?}",
        prog.data_items.len(), data_pos, ws_pos, &token_kinds[..20.min(token_kinds.len())]);
}

#[test]
fn parse_data_items_no_trailing_newline() {
    // Same source but without leading newline
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. CALC.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-MSG PIC X(20) VALUE \"From COBOL\".\nPROCEDURE DIVISION.\n    DISPLAY WS-MSG.\n    STOP RUN.";
    let prog = vybe_parser_cobol::parse(src).expect("parse failed");
    assert!(!prog.data_items.is_empty(),
        "Should have data items, got {} items", prog.data_items.len());
}

#[test]
fn parse_data_items_with_env_div() {
    let src = r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CALC.
ENVIRONMENT DIVISION.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MSG PIC X(20) VALUE "From COBOL".
PROCEDURE DIVISION.
    DISPLAY WS-MSG.
    STOP RUN.
"#;
    let prog = vybe_parser_cobol::parse(src).expect("parse failed");
    assert!(!prog.data_items.is_empty(),
        "Should have data items with ENV DIV, got {} items", prog.data_items.len());
}

#[test]
fn parse_data_items_with_special_names() {
    let src = r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CALC.
ENVIRONMENT DIVISION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MSG PIC X(20) VALUE "From COBOL".
PROCEDURE DIVISION.
    DISPLAY WS-MSG.
    STOP RUN.
"#;
    let prog = vybe_parser_cobol::parse(src).expect("parse failed");
    assert!(!prog.data_items.is_empty(),
        "Should have data items with SPECIAL-NAMES, got {} items", prog.data_items.len());
}
