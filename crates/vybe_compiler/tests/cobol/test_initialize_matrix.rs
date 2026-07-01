use super::helpers::compile_ok_check;

fn make_program(replacing: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-GROUP.\n   05 WS-A PIC 9(3) VALUE 123.\n   05 WS-B PIC X(5) VALUE \"HELLO\".\nPROCEDURE DIVISION.\n    INITIALIZE WS-GROUP {replacing}.\n    STOP RUN."
    )
}

#[test]
fn test_initialize_replacement_matrix() {
    let variants = [
        "",
        "REPLACING NUMERIC BY 5",
        "REPLACING ALPHANUMERIC BY \"WORLD\"",
        "REPLACING NUMERIC DATA BY 9",
        "REPLACING ALPHANUMERIC DATA BY \"X\"",
        "REPLACING NUMERIC BY 0 ALPHANUMERIC BY \"Z\"",
        "REPLACING NUMERIC BY 7",
        "REPLACING ALPHANUMERIC BY \"Q\"",
    ];

    for variant in variants {
        assert!(compile_ok_check(&make_program(variant)), "initialize failed for {variant}");
    }
}
