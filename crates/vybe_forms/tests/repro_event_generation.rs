use vybe_forms::EventType;

#[test]
fn test_event_signature_generation() {
    let evt = EventType::Click;
    let params = evt.parameters();
    // All events use "sender As Object, e As EventArgs" as base signature
    assert_eq!(params, "sender As Object, e As EventArgs");
    
    let evt_mouse = EventType::MouseDown;
    let params_mouse = evt_mouse.parameters();
    assert!(params_mouse.contains("MouseEventArgs"));
}
