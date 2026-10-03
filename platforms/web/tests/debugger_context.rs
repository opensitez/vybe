use vybe_platform_web::html;

#[test]
fn debugger_document_does_not_turn_console_program_into_window() {
    let id = html::active_document_for_debugger();
    assert!(html::has_browsing_context());
    assert!(!html::has_guest_browsing_context());

    assert_eq!(html::active_document(), id);
    assert!(html::has_guest_browsing_context());
}
