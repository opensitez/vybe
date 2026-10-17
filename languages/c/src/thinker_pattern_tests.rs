#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn embedded_door_thinker_callback_preserves_registration_and_derived_updates() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/embedded_door_thinker_callback.c")]);
}
