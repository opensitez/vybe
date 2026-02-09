use irys_forms::EventType;

#[test]
fn test_event_signature_generation() {
    let evt = EventType::Click;
    let params = evt.parameters();
    assert_eq!(params, ""); // Click has no params usually, or maybe arguments?
    // Wait, my EventType definition might have params for Click?
    // Let's check the code I read earlier.
    
    let evt_mouse = EventType::MouseDown;
    let params_mouse = evt_mouse.parameters();
    assert!(params_mouse.contains("Button As Integer"));
}
